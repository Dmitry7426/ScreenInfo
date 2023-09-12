# Import the module.
 Import-Module PSExcel  

# Опрос файла инициализации и формирование словаря, временного файла

$slovar = @(Get-Content -Path c:\ScreenInfo\conf.ini)
$f0,$dir_patch  = $slovar[0].split('>')
$f1,$name_in_file  = $slovar[1].split('>')
$f2,$name_out_file  = $slovar[2].split('>')
$FullPatchIn = "$name_in_file" -replace ' ', ''
$FullPatchOut = "$dir_patch$name_out_file" -replace ' ', ''
$TempFile = New-TemporaryFile

# Первоначальное заполнение HTML либо создание если отсутствует страница

$Len = Get-ChildItem -Path $FullPatchIn | Select-Object LastWriteTime # Первоначальная дата модификации файла Excel
    
# Функция опроса, формирования временного файла
function Creat_Data_Html
    {
        Clear-Content $TempFile
        $ExcelPackage = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $FullPatchIn
        $ExcelWorkSheet = $ExcelPackage.Workbook.Worksheets['INFO']
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
        $total_service = $ExcelWorkSheet.Cells[("G49")].Value

        # Стилизация отображения тэгов

        $st_risk = "<img hidden src='c:\ScreenInfo\image\arrow.png' width='20' height='10' alt='Иллюстрация'>"
        $sm_count = 0

        #  Переборка смен и людей в смене

        for (;$a -le 48; $a++)
            {
                if ($ExcelWorkSheet.Cells[("A$a")].Value -gt 0) 
                    {
                        $sm_fio += "<div class=alg><a>" + @($ExcelWorkSheet.Cells[("A$a")].Value) + "</a>" + "<a>" + @($ExcelWorkSheet.Cells[("B$a")].Value) + "</a></div>"
                        $sm_stat += "<a>" + @($ExcelWorkSheet.Cells[("B$a")].Value) + "</a>"
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
                        $of_fio += "<div class=alg><a>" + @($ExcelWorkSheet.Cells[("C$a")].Value) + "</a>" + "<a>" + @($ExcelWorkSheet.Cells[("D$a")].Value) + "</a></div>"
                        $of_stat += "<a class='status_office'>" + @($ExcelWorkSheet.Cells[("D$a")].Value) + "</a>"     
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
                        $guests += "<div class='alg oxr'><a>" + @($ExcelWorkSheet.Cells[("E$a")].Value) + "</a>" + "<a>" + @($ExcelWorkSheet.Cells[("F$a")].Value) + "</a></div>"
                    }
                else 
                    {
                        $guests += @("")
                    }  

        # Переборка Охрана/сервсис
      
              $oxr += "<div class='alg oxr'><a>" + @($ExcelWorkSheet.Cells[("g$a")].Value + "</a></div>")  
            }

        # Переборка Рисков

        $rsk = 2
        $ExcelWorkSheetRisk = $ExcelPackage.Workbook.Worksheets['Риски']

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
                        <div id='curr' class='curr'></div>
                         <div class='status'>
                            <div class='st1'>
                                <div class=alg><h2>Фамилия</h2><h2>Статус</h2></div>
                                $sm_fio
                                <a class='total_ch'>Всего: $total_smena</a>
                            </div>

                            

                            <div class='st1'>
                                <div class=alg><h2>Фамилия</h2><h2>Статус</h2></div>
                                $of_fio
                                <a class='total_ch'>Всего: $total_office</a>
                            </div>
                         
                            
                                
                            
                         
                            <div class='st1'>
                                <div class=alg><h2>К кому</h2><h2>Кол-во</h2></div>
                                $guests
                                <a class='total_ch'>Всего: $total_guest</a>

                            </div>
                         
                            <div class='st1'>
                                <div class=alg><h2>Охрана/Сервис</h2></div>
                                $oxr
                                <a class='total_ch'>Всего: $total_service</a>
                            </div>

                                <div class='riski'>
                                    <a class='borders'>Повышенный риск</a>
                                    <div>
                                        $sm_risk
                                        
                                    </div>
                                    <br>
                                    <a class='borders'>Возможные риски</a>
                                    <div>
                                        $sm_risk_style
                                    </div>

                                    <br>
                                    <a class='borders'>Всего на предприятии</a> 
                                    <br>
                                    <div class='total'>
                                        <h1>$peoples</h1>
                                    </div>
                                    <br>
                                    <div class='totals'></div>
                               
                            
                            <a></a>
                            
                            <a></a>
                            
                            <a></a>

                      
                    
                </div>
            </body>
            </html>"

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

function G_My 
    {
        Write-Host $Len.LastWriteTime
        $Len2 = Get-ChildItem -Path $FullPatchIn | Select-Object LastWriteTime
        Write-Host $Len2.LastWriteTime
        if ($Len.LastWriteTime -ne $Len2.LastWriteTime) 
            {
                Creat_Data_Html
                Write-Host "Source file modify! Refresh HTML page!"
                $Len = $Len2       
                Copy-Item -Path $TempFile -Destination $FullPatchOut 
                Remove-Item $TempFile.FullName   
                G_My               
            }
        else 
            {
                Write-Host "File not modify! No refresh HTML page!"
            }
        Start-Sleep -Seconds 60
        G_My
    }

# Функция открытия информационного экрана

function Html_Load  
    {
        Start-Process microsoft-edge:$FullPatchOut
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($FullPatchOut)
        Start-Sleep 1
        #$Output = $wshell.Popup("$FullPachOut")
        Start-Process $FullPatchOut
        $wshell.SendKeys('{F11}')
    }


   Creat_Data_Html

   Html_Load

   G_My