

# Import the module.
 Import-Module PSExcel  
 

# Опрос файла инициализации и формирование словаря, временного файла

$slovar = @(Get-Content -Path C:\ScreenInfo\conf.ini)
$f0,$dir_patch  = $slovar[0].split('>')
$f1,$name_in_file  = $slovar[1].split('>')
$f2,$name_out_file  = $slovar[2].split('>')
$FullPatchIn = "$name_in_file" -replace ' ', ''
$FullPatchOut = "$dir_patch$name_out_file" -replace ' ', ''
$TempFile = New-TemporaryFile

Write-Host $FullPatchIn

# Первоначальное заполнение HTML либо создание если отсутствует страница

$Len = Get-ChildItem -Path $FullPatchIn | Select LastWriteTime # Первоначальная дата модификации файла Excel
    
# Функция опроса, формирования временного файла
function Creat-Data-Html
    {
        Clear-Content $TempFile
       
        $ExcelPackage = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $FullPatchIn

        # Select First Worksheet
        $ExcelWorkSheet = $ExcelPackage.Workbook.Worksheets['INFO']

        #$Worksheet.Cells[('A6')].Value

       
        #$ExcelObj = New-Object -comobject Excel.Application
        #$ExcelObj.Visible = $false
        #$ExcelWorkBook = $ExcelObj.Workbooks._Open($FullPatchIn)
        #$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Info")
        $dt = Get-Date -DisplayHint Date
        $a = 4
        $sm_fio = @()
        $sm_stat = @()
        $sm_risk = @()
        $sm_risk_style = @()
        $te = @()
        $oxr = @()
        $total_smena = $ExcelWorkSheet.Cells[("B49")].Value
        $total_office = $ExcelWorkSheet.Cells[("D49")].Value
        $total_guest = $ExcelWorkSheet.Cells[("F49")].Value

        # Стилизация отображения тэгов

        $st_risk = "<img src='c:\ScreenInfo\image\arrow.png' width='20' height='10' alt='Иллюстрация'>"
        $sm_count = 0

        #  Переборка смен и людей в смене

        for (;$a -le 48; $a++)
            {
                if ($ExcelWorkSheet.Cells[("A$a")].Value -gt 0) 
                    {
                        $sm_fio += @($ExcelWorkSheet.Cells[("A$a")].Value) + "<br>"
                        $sm_stat += @($ExcelWorkSheet.Cells[("B$a")].Value) + "<br>"
                        $sm_count += 1       
                    }
                else 
                    {
                        $sm_fio += @("") + "<br>"
                        $sm_stat += @("") + "<br>"
                    }

        # Переборка офис
      
                if ($ExcelWorkSheet.Cells[("C$a")].Value -gt 0) 
                    {
                        $of_fio += @($ExcelWorkSheet.Cells[("C$a")].Value) + "<br>"
                        $of_stat += @($ExcelWorkSheet.Cells[("D$a")].Value) + "<br>"     
                        $of_count += 1
                    }
                else 
                    {
                        $of_fio += @("") + "<br>"
                        $of_stat += @("") + "<br>"
                    }

        # Переборка посетителей

                if ($ExcelWorkSheet.Cells[("F$a")].Value -ne 0) 
                    {
                        $guests += @($ExcelWorkSheet.Cells[("E$a")].Value + " - " + $ExcelWorkSheet.Cells[("F$a")].Value + "<br>")
                    }
                else 
                    {
                        $guests += @("")
                    }  

        # Переборка Охрана/сервсис
      
              $oxr += @($ExcelWorkSheet.Cells[("g$a")].Value + "<br>")  

            }

        # Переборка Рисков

        $rsk = 2
        $ExcelWorkSheetRisk = $ExcelPackage.Workbook.Worksheets['Риски']
        #$ExcelWorkSheetRisk = $ExcelWorkBook.Sheets.Item("Риски")

        #Write-Host $oxr

        for (;$rsk -le 10; $rsk++) 
            {
                if ($ExcelWorkSheetRisk.Cells[("A$rsk")].Value -eq 1) 
                    {
                        $sm_risk += "<p class=risk_up>" + "$st_risk" +  @($ExcelWorkSheetRisk.Cells[("B$rsk")].Value) + "</p>"
                    }
                else 
                    {
                        $sm_risk_style += "<p class=risk_down>" + "$st_risk" +  @($ExcelWorkSheetRisk.Cells[("B$rsk")].Value) + "</p>"
                    }
            }

        # Людей, находящихся на заводе в текущий момент

        $peoples = $ExcelWorkSheet.Cells[("h28")].Value 

        # Формирование структуры файла HTML во временном файле

        Add-Content -Encoding UTF8 $TempFile "
            <!DOCTYPE html>
            <html lang='ru'>
            <head>
                <meta charset='UTF-8'>
                <meta name='viewport' content='width=device-width, initial-scale=1.0'>
                <link rel='stylesheet' href='./style/style.css'>
                <script type='text/javascript' src='./scrypt/scrypt.js'></script>
                <title>Informations Screen</title>
            </head>
            <body>
                <div class='backgr'>
                    <div id='curr' class='curr'></div>          
                    <div class='mytable'>
                        <table> 
                            <tr>
                                <th>Фамилия</th>
                                <th>Статус</th>
                                <th>Фамилия</th>
                                <th>Статус</th>
                                <th>К кому - кол-во</th>
                                <th>Охрана/Сервис</th>
                            </tr>
                            <tr>
                                <td>$sm_fio  </td>
                                <td>$sm_stat </td>
                                <td>$of_fio</td>
                                <td>$of_stat</td>
                                <td>$guests</td>
                                <td>$oxr</td>
                                <td class='riski'>
                                    <a class='borders'>Повышенный риск</a>
                                    <div>
                                        $sm_risk
                                        
                                    </div>
                                    <br>
                                    <a class='borders'>Возможные риски</a>
                                    <p class='risk_down'>$sm_risk_style</p>
                                    <br>
                                    <a class='borders'>Всего на предприятии</a> 
                                    <br>
                                    <div class='total'>
                                        <a>$peoples</a>
                                    </div>
                                    <br>
                                    <div class='totals'></div>
                                </td>
                            </tr>
                            <td>Всего: $total_smena</td>
                            <td></td>
                            <td>Всего: $total_office</td>
                            <td></td>
                            <td>Всего: $total_guest</td>
                            <td></td>

                        </table>
                    </div>
                </div>
            </body>
            </html>"

        #$ExcelObj.ActiveWindow.Close($false) # Закрываем активное окно Excel
        #$ExcelObj.Quit() # Завершение работы приложения

        # Проверка наличия файла на старте и в случает отсутсвия создание

        if (Test-Path -Path $FullPatchOut) 
            { 
                Write-Host "HTML Files Found! Go!"
            }
        else { 
                Write-Host "File not found. Create Files"
                Copy-Item -Path $TempFile -Destination $FullPatchOut
             }      

    }  # Конец функции Creat-Data-Html

 # Открываем браузер на весь экран с информационным экраном
 # Стартуем в цикл функцию контроля за изменением исходного файла

# Функция Проверка изменения информационного файла

function G-My 
    {
        Write-Host $Len.LastWriteTime
        $Len2 = Get-ChildItem -Path $FullPatchIn | Select LastWriteTime
        Write-Host $Len2.LastWriteTime
        if ($Len.LastWriteTime -ne $Len2.LastWriteTime) 
            {
                Creat-Data-Html
                Write-Host "Source file modify! Refresh HTML page!"
                $Len = $Len2       
                Copy-Item -Path $TempFile -Destination $FullPatchOut 
                Remove-Item $TempFile.FullName   
                G-My               
            }
        else 
            {
                Write-Host "File not modify! No refresh HTML page!"
            }

        Start-Sleep -Seconds 60
        G-My
    }

# Функция открытия информационного экрана

function Html-Load  
    {
        start microsoft-edge:$FullPatchOut
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($FullPatchOut)
        Sleep 1
        #$Output = $wshell.Popup("$FullPachOut")
        start $FullPatchOut
        $wshell.SendKeys('{F11}')
    }

function T-Echo 
    {
        $Output = $wshell.Popup("$FullPachOut")
    }

function Refresh-WebPages 
    {        
        Where-Object { $_.Document.url -like 'file:///C:/intel/1.html' } | 
        ForEach-Object { $_.Refresh() }
    }

   Creat-Data-Html

   Html-Load

   G-My