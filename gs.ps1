$infolder = 'c:\Users\mroli\Documents\WHB\Werkstatthandbuch J9_in\'
$outfolder = 'c:\Users\mroli\Documents\WHB\Werkstatthandbuch J9_out\'
$outfilename = 'Werkstatthandbuch j9.pdf'
$gspath = 'C:\Program Files\gs\gs9.53.3\bin\gswin64c.exe'
$gspath2 = 'C:\Program Files\gs\gs9.50\bin\gswin64c.exe'
$tesseracttraineddata = 'C:/Program Files/gs/tessdata/'
$tesseractlanguage = 'deu'
$gsparameters1 = '-o "{CURRENTFILE}.txt" -sDEVICE=ink_cov "{CURRENTFILE}"'
$gsparameters2 = '-o "{CLEANFILE}" -sDEVICE=pdfwrite -sPageList={PAGENUMBER} "{CURRENTFILE}"'
$gsparameters3 = '-o "{OUTFILE}" -sDEVICE=pdfwrite pdfmarks '
$gsparameters4 = '--permit-file-read="{TESSDATA}" -sDEVICE=pdfocr8 -sOCRLanguage={TESSLANG} -o "{OUTFILE}" -r600 -dDownscaleFactor=3 "{CURRENTFILE}"' ##
##$gsparameters5 = '-dBatch -dNOPAUSE -o "{OUTFILE}" -sDEVICE=pdfwrite -dDOPDFMARKS {PDFMARK} {CURRENTFILE}'
$inklimit = 0.1 ## to remove empty pages

$pdfmark = '[ /Title (Document title)
/Author (Author name)
/Subject (Subject description)
/Keywords (comma, separated, keywords)
/Creator (application name or creator note)
/Producer (PDF producer name or note)
/DOCINFO pdfmark'

##$pdfmark = '[ /Title (Document title) /Author (Author name) /DOCINFO pdfmark'

$env:TESSDATA_PREFIX = $tesseracttraineddata

$pdfmark=$pdfmark.replace('(Document title)','('+$outfilename+')')
$pdfmark=$pdfmark.replace('(Author name)','(Ghostscript)')
$pdfmark=$pdfmark.replace('(Subject description)','(Alles für den J, Alles für den Club!)')
$pdfmark=$pdfmark.replace('(application name or creator note)','(Ghostscript)')
$pdfmark=$pdfmark.replace('(PDF producer name or note)','(Ghostscript)')
$pdfmark=$pdfmark.replace('(comma, separated, keywords)','(peugeotj9,j9,whb,werkstatthandbuch)')

$step1 = $false ##scan for ink
$step2 = $false ##analize ink data
$step3 = $true ## remove blank pages with help of the analysed data and create pdfmarks
$step4 = $false ## OCR files
$step5 = $true ## merge files and add toc


##region scan pdfs for ink
Write-host (get-date).ToString('s'): Starting Step 1
if ($step1 -eq $true) {
    foreach ($file in (gci ($infolder+'*') -Include *_OCR.pdf <#-Exclude *OCR.pdf#> | sort Name)) {
        Write-host (get-date).ToString('s'): Scanning $file.Name
        $gsparam = $null
        $gsparam = $gsparameters1.Replace('{CURRENTFILE}',$file.FullName)
        Start-Process -FilePath $gspath -ArgumentList $gsparam -Wait
        Write-host (get-date).ToString('s'): Finished $file.Name
    }
}

## check ink files
Write-host (get-date).ToString('s'): Starting Step 2
if ($step2 -eq $true) {
    foreach ($file in (gci ($infolder+'*') -Include *.pdf.txt | sort Name)) {
        Write-host (get-date).ToString('s'): checking $file.Name
        $inkcov = Get-content $file
        $pagestokeep = New-Object -TypeName 'System.Collections.ArrayList'
        $i = 0
        foreach ($line in $inkcov) {
            $i++
            if ($line.split(' ')[-3] -gt $inklimit) {
                $pagestokeep.Add($i) >> $null
                Write-host $file +'|'+ $i +'|'+ $line.split(' ')[-3] +'|'+ 'ok'
            } else {
                Write-host $file +'|'+ $i +'|'+ $line.split(' ')[-3] +'|'+ 'blank'
            }
        }
        Write-host $pagestokeep -join(',')
        $pagestokeep -join(',') > ($file.fullname + '.csv')
    }
}

## create cleaned files and pdfmark
Write-host (get-date).ToString('s'): Starting Step 3
if ($step3 -eq $true) {
    $i = 1
    foreach ($file in (gci ($infolder+'*') -Include *_OCR.pdf <#-Exclude *OCR.pdf#> | sort Name)) {
        Write-host (get-date).ToString('s'): Checking $file.Name
        if(Test-Path($file.fullname+'.txt.csv')) {
            $pagelist = $null
            write-host $file.fullname
            $pagelist = (Get-content ($file.fullname+'.txt.csv'))
            $gsparam = $gsparameters2.replace('{CURRENTFILE}',$file.FullName).replace('{CLEANFILE}',$outfolder+$file.Name+'clean.pdf').replace('{PAGENUMBER}',$pagelist)
            Write-host $gsparam
            Start-Process -FilePath $gspath2 -ArgumentList $gsparam -Wait -NoNewWindow
            $pdfmark = $pdfmark + "`n[ /Title ({Title Page}) /Count 0 /Page {page} /OUT pdfmark".replace('{Title Page}',(($file.Name).split('.')[0])).replace('{page}',$i)
            Write-host $i ($pagelist.split(',')).count ($i + ($pagelist.split(',')).count)
            $i = ($i + ($pagelist.split(',')).count)

        } else {Write-host $file.fullname+'.txt.csv' not found}
    }
    $pdfmark | Out-File -FilePath ($outfolder+'pdfmarks') -Encoding ascii
    ##$pdfmark | Out-File -FilePath ($outfolder+'pdfmarks') -Encoding unicode
}
##ocr them
Write-host (get-date).ToString('s'): Starting Step 4
if ($step4 -eq $true) {
    foreach ($file in (gci ($outfolder+'*') -Include *.pdfclean.pdf <#-Exclude *_OCR.pdfclean.pdf#> | sort Name)) {
        Write-host (get-date).ToString('s'): checking $file.Name
        $gsparam = $gsparameters4.replace('{TESSDATA}',$tesseracttraineddata).replace('{TESSLANG}',$tesseractlanguage).replace('{OUTFILE}',$outfolder+$file.Name+'OCR.pdf').replace('{CURRENTFILE}',$File.Fullname)
        Start-Process -FilePath $gspath -ArgumentList $gsparam -Wait -NoNewWindow
    }
}

Write-host (get-date).ToString('s'): Starting Step 5
if ($step5 -eq $true) {
    $filelist = (gci ($outfolder+'*') -Include *.pdfclean.pdf <#-Exclude *OCR.pdfclean.pdf#> | sort Name)
    $gsparam = $gsparameters3.replace('{OUTFILE}',$outfilename) + ' ' + (($filelist | Select Name).Name -join (' '))
    Start-Process -FilePath $gspath -ArgumentList $gsparam -WorkingDirectory $outfolder -Wait -NoNewWindow
}
