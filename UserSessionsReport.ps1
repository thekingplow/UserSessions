#v1.4 (18.07.2025)
#Developed by Danilovich M.D.




param(
    [string]$ComputerName
)

# Установка кодировки на UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Путь к файлу с черным списком компьютеров
$blacklistFile = Join-Path -Path $PSScriptRoot -ChildPath "blacklist.txt"


$logfile = Join-Path -Path $PSScriptRoot -ChildPath "logs\session_log.txt"










# Начало записи транскрипта (для лога)
Start-Transcript -Path $logfile


# Чтение черного списка компьютеров
try {
    if (Test-Path $blacklistFile) {
        $blacklist = Get-Content -Path $blacklistFile
        Write-Host "Черный список загружен:" -ForegroundColor Green
        $blacklist | ForEach-Object { Write-Host $_  -ForegroundColor DarkGray}
    } else {
        $blacklist = @()
        Write-Host "Файл с черным списком не найден. Продолжаем без исключений." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Не удалось прочитать файл с черным списком." -ForegroundColor Red
    #exit
}





# Получение списка компьютеров из Active Directory, исключая OU "Computers"
try {
    Import-Module "ActiveDirectory"

    if ($ComputerName) {
        # Если указан конкретный компьютер, проверяем только его
        $computersListFromAd = @($ComputerName)
    } else {
        # Получение всех компьютеров из AD
        $computersListFromAd = Get-ADComputer -Filter * -Properties dNSHostName |
                               Where-Object { $_.DistinguishedName -notlike "*OU=Computers,*" } |
                               Select-Object -ExpandProperty dNSHostName | ForEach-Object { $_ -replace "\.domain\.local$", "" }
    }

    if ($computersListFromAd.Count -eq 0) {
        Write-Host "Не удалось получить компьютеры из Active Directory." -ForegroundColor Red
        exit
    } else {
        # Записываем в лог полную информацию о полученных компьютерах
        Write-Host "Полученные компьютеры:" -ForegroundColor Green
        $computersListFromAd | ForEach-Object { Write-Host $_ }
    }

} catch {
    Write-Host "Не удалось подключиться к Active Directory." -ForegroundColor Red
    exit
}





# Исключение компьютеров из черного списка
$computersToCheck = $computersListFromAd | Where-Object { $_ -notin $blacklist }

# Записываем в лог полную информацию о компьютерах для проверки
Write-Host "Компьютеры для проверки после исключения черного списка:" -ForegroundColor Green
$computersToCheck | ForEach-Object { Write-Host $_ }

# Создаем пустой массив для хранения результатов
$allUsers = @()

foreach ($computer in $computersToCheck) {
    try {
        # Создаем сессию PowerShell для указанного компьютера с использованием текущего пользователя (Kerberos)
        $s = New-PSSession -ComputerName $computer -Authentication Kerberos -ErrorAction Stop

        # Запускаем скрипт на удаленной машине
        $Users = Invoke-Command -Session $s -ScriptBlock {
            $Computer = $env:COMPUTERNAME

            #Возвращает информацию о пользовательских сеансах 
            $Users = query user /server:$Computer 2>&1

            
            $Users = $Users | ForEach-Object {
                try {
                    (($_ -replace ">" -replace "(?m)^([A-Za-z0-9]{3,})\s+(\d{1,2}\s+\w+)", '$1  none  $2' -replace "\s{2,}", "," -replace "none", $null).Trim())
                } catch {
                    Write-Host "Ошибка при обработке строки: $_" -ForegroundColor Red
                }
            } | ConvertFrom-Csv 

            if ($Users.Count -eq 0) {
                Write-Host "На компьютере $Computer нет пользователей." -ForegroundColor Yellow
                return $null
            } else {

            #Формирование объектов результата сессии
                foreach ($User in $Users) {
                    [PSCustomObject]@{
                        ComputerName = $Computer
                        Username = $User.USERNAME
                        SessionState = $User.STATE.Replace("Disc", "Disconnected")
                    }
                }
            }
        }

        # Добавляем результаты только если были найдены пользователи
        if ($Users -ne $null) {
            $allUsers += $Users | Select-Object ComputerName, Username, SessionState
        }

    } catch {
        Write-Host "Произошла ошибка на компьютере $($computer): $_" -ForegroundColor Red
    }
}

# Завершение записи транскрипта для включения логов в файл
Stop-Transcript










# Удаление дубликатов по ComputerName, Username и SessionState
$allUsers = $allUsers | Sort-Object ComputerName, Username, SessionState -Unique


# Если $allUsers пустой, значит на всех компьютерах не было пользователей
if ($allUsers.Count -eq 0) {
    Write-Host "На всех компьютерах нет пользователей."
} else {

    # Формируем тело письма
    $body = $allUsers | ConvertTo-Html | Out-String
}

# Чтение логов для отправки по почте
$logs = Get-Content -Path $logfile | Out-String







# Убираем ненужные части из логов
$logsWithoutComputers = $logs -replace "(?s)(Полученные компьютеры:(.|\r?\n)*?)(?=INFO)", "" `
                                 -replace "(?s)(Компьютеры для проверки после исключения черного списка:(.|\r?\n)*?)(?=INFO)", ""
# Разбиваем текст на строки
$logLines = $logsWithoutComputers -split "`r?`n"

# Выбираем строки с "INFO", убираем дубликаты и сортируем их
$infoLines = $logLines | Where-Object { $_ -match "^INFO:" } | Sort-Object -Unique

# Оставляем остальные строки без изменений
$otherLines = $logLines | Where-Object { $_ -notmatch "^INFO:" }

# Собираем итоговый текст: несортированные строки и отсортированные INFO без дубликатов
$sortedLogs = @($otherLines + $infoLines) -join "`r`n"

# Результат
$logsWithoutComputers = $sortedLogs








# Заголовок и стиль HTML-таблицы
$body = @"
<html>
<head>
<style>
  table { border-collapse: collapse;  }
  th, td { border: 1px solid #ddd; padding: 8px; font-family: Arial; width: 30%; }
  th { background-color: #f2f2f2; }
</style>
</head>
<body>
<h3>Активные пользовательские сессии</h3>
<table>
  <tr>
    <th>Имя компьютера</th>
    <th>Имя пользователя</th>
    <th>Время сессии</th>
  </tr>
"@

# Добавление строк с результатами
foreach ($user in $allUsers) {
    $body += "<tr><td>$($user.ComputerName)</td><td>$($user.Username)</td><td>$($user.SessionState)</td></tr>`n"
}



# Закрытие таблицы и добавление логов
$body += @"
</table>
<br><br><strong>Логи выполнения:</strong><br><pre>$logsWithoutComputers</pre>
</body>
</html>
"@







# Добавляем очищенные логи в тело письма
$body += "<br><br><strong>Логи выполнения:</strong><br><pre>$logsWithoutComputers</pre>"

    # Здесь укажите параметры для отправки письма
    $smtpServer = "mail.server.com"
    $fromAddress = "from@mail.com"  #Указывается почта от кого письмо будет приходить
    #$toAddress = "admin.support@itsbel.by"  #Указывается почта кому письмо будет приходить
    $toAddress = "to@mail.com"  #Указывается почта кому письмо будет приходить
    $subject = "Активные сессии"

# Отправляем письмо с указанием кодировки UTF-8 для тела письма
Send-MailMessage -SmtpServer $smtpServer -From $fromAddress -To $toAddress -Subject $subject -Body $body -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
