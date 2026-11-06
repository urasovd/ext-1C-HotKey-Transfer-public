#Requires -Version 7
# Integration test for WS_HotKeyTransfer. Run ONLY against test infobases: it creates catalog items
# and rewrites one document's movements and one information-register record on the target.
param(
    [Parameter(Mandatory = $true)] [string] $Server,
    [Parameter(Mandatory = $true)] [string] $SourceRef,
    [Parameter(Mandatory = $true)] [string] $TargetRef,
    [Parameter(Mandatory = $true)] [string] $TargetUrl,
    [Parameter(Mandatory = $true)] [string] $User,
    [Parameter(Mandatory = $true)] [string] $PasswordFile
)

$ErrorActionPreference = 'Stop'
$password = (Get-Content -LiteralPath $PasswordFile -Raw -Encoding UTF8).Trim()
$connector = New-Object -ComObject V83.COMConnector
$results = [System.Collections.Generic.List[object]]::new()
$flags = [Reflection.BindingFlags]

function G($o, $n) { , [System.__ComObject].InvokeMember($n, $flags::GetProperty, $null, $o, $null) }
function S($o, $n, $v) { [void][System.__ComObject].InvokeMember($n, $flags::SetProperty, $null, $o, @(, $v)) }
function M($o, $n, [object[]] $a = @()) {
    # PowerShell wraps values in PSObject, and COM rejects such wrappers as arguments.
    $plain = [object[]]::new($a.Count)
    for ($i = 0; $i -lt $a.Count; $i++) { if ($null -ne $a[$i]) { $plain[$i] = $a[$i].PSObject.BaseObject } }
    , [System.__ComObject].InvokeMember($n, $flags::InvokeMethod, $null, $o, $plain)
}

function Connect-Base([string] $ref) {
    , $connector.Connect("Srvr=`"$Server`";Ref=`"$ref`";Usr=`"$User`";Pwd=`"$password`"")
}

function New-Selection($v, [object[]] $values) {
    $list = M $v 'NewObject' @('Массив')
    foreach ($value in $values) { [void](M $list 'Добавить' @(, $value)) }
    , $list
}

function Get-Items($collection) {
    $count = M $collection 'Количество'
    for ($i = 0; $i -lt $count; $i++) { , (M $collection 'Получить' @($i)) }
}

function Add-Result([string] $test, [bool] $passed, [string] $detail = '') {
    $results.Add([pscustomobject]@{ Test = $test; Passed = $passed; Detail = ($detail.Replace($password, '***')) })
}

function Invoke-Test([string] $test, [scriptblock] $body) {
    try { & $body } catch {
        $message = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
        Add-Result $test $false "$message @ $(($_.ScriptStackTrace -split "`n")[0..1] -join ' < ')"
    }
}

function Invoke-Check($v, $selection) {
    , (M (G $v 'WS_ТранспортСервер') 'Проверить' @($selection, $TargetUrl, ''))
}

function Invoke-Transfer($v, $selection, [string] $reason = '', $unchecked = $null) {
    if ($null -eq $unchecked) { $unchecked = M $v 'NewObject' @('Массив') }
    , (M (G $v 'WS_ТранспортСервер') 'Перенести' @($selection, $TargetUrl, '', $reason, $unchecked))
}

function Get-Rows($check) { , @(Get-Items (G $check 'Строки')) }

function Get-RefByUuid($v, $managerName, $objectName, [string] $uuid) {
    $manager = G (G $v $managerName) $objectName
    , (M $manager 'ПолучитьСсылку' @((M $v 'NewObject' @('УникальныйИдентификатор', $uuid))))
}

function Test-Exists($ref) { $null -ne (M $ref 'ПолучитьОбъект') }

function Get-Uuid($v, $ref) { [string](M $v 'XMLСтрока' @(, $ref)) }

# Finds a hierarchical catalog where a parent and a child can be created, so the child
# has a reference the target does not have yet. Nothing here depends on the configuration.
function New-LinkedCatalogItems($v) {
    $suffix = Get-Date -Format 'yyyyMMdd_HHmmss'
    foreach ($catalog in @(Get-Items (G (G $v 'Метаданные') 'Справочники'))) {
        if (-not (G $catalog 'Иерархический')) { continue }
        if ((M (G $catalog 'Владельцы') 'Количество') -gt 0) { continue }
        $name = G $catalog 'Имя'
        $manager = G (G $v 'Справочники') $name
        try {
            $parent = M $manager 'СоздатьГруппу'
            S (G $parent 'ОбменДанными') 'Загрузка' $true
            S $parent 'Наименование' "HKT_PARENT_$suffix"
            [void](M $parent 'Записать')
            $child = M $manager 'СоздатьЭлемент'
            S (G $child 'ОбменДанными') 'Загрузка' $true
            S $child 'Наименование' "HKT_CHILD_$suffix"
            S $child 'Родитель' (G $parent 'Ссылка')
            [void](M $child 'Записать')
            return [pscustomobject]@{ Catalog = $name; Parent = (G $parent 'Ссылка'); Child = (G $child 'Ссылка') }
        } catch { continue }
    }
    throw 'No suitable hierarchical catalog found.'
}

function Find-DocumentWithMovements($v) {
    foreach ($document in @(Get-Items (G (G $v 'Метаданные') 'Документы'))) {
        $records = G $document 'Движения'
        if ((M $records 'Количество') -eq 0) { continue }
        $name = G $document 'Имя'
        $query = M $v 'NewObject' @('Запрос', "ВЫБРАТЬ ПЕРВЫЕ 20 Т.Ссылка КАК Ссылка ИЗ Документ.$name КАК Т ГДЕ Т.Проведен")
        $selection = M (M $query 'Выполнить') 'Выбрать'
        while (M $selection 'Следующий') {
            $ref = G $selection 'Ссылка'
            $movements = G (M $ref 'ПолучитьОбъект') 'Движения'
            foreach ($register in $records) {
                $registerName = G $register 'Имя'
                $set = G $movements $registerName
                [void](M $set 'Прочитать')
                $count = M $set 'Количество'
                if ($count -gt 0) {
                    return [pscustomobject]@{ Document = $name; Uuid = (Get-Uuid $v $ref); Ref = $ref; Register = $registerName; Count = $count }
                }
            }
        }
    }
    throw 'No posted document with movements found.'
}

# Independent register with a record whose whole key can be rebuilt from a query row.
function Find-RegisterRecord($v) {
    foreach ($register in @(Get-Items (G (G $v 'Метаданные') 'РегистрыСведений'))) {
        $name = G $register 'Имя'
        $dimensions = @(Get-Items (G $register 'Измерения') | ForEach-Object { G $_ 'Имя' })
        if ($dimensions.Count -eq 0) { continue }
        $query = M $v 'NewObject' @('Запрос', "ВЫБРАТЬ ПЕРВЫЕ 1 * ИЗ РегистрСведений.$name")
        $queryResult = M $query 'Выполнить'
        $columns = G $queryResult 'Колонки'
        if ($null -ne (M $columns 'Найти' @('Регистратор'))) { continue }
        $row = M $queryResult 'Выбрать'
        if (-not (M $row 'Следующий')) { continue }
        $key = @{}
        $supported = $true
        foreach ($dimension in $dimensions) {
            $value = G $row $dimension
            $key[$dimension] = $value
            if ($value -is [System.__ComObject]) {
                try { $supported = $supported -and ([string](M (M $value 'Метаданные') 'ПолноеИмя')) -match '^(Справочник|Документ)\.' } catch { $supported = $false }
            }
        }
        if (-not $supported) { continue }
        if ($null -ne (M $columns 'Найти' @('Период'))) { $key['Период'] = G $row 'Период' }
        return [pscustomobject]@{ Register = $name; Key = $key }
    }
    throw 'No independent register with records found.'
}

function New-Filtered-Set($v, [string] $register, [hashtable] $key) {
    $set = M (G (G $v 'РегистрыСведений') $register) 'СоздатьНаборЗаписей'
    $filter = G $set 'Отбор'
    foreach ($field in $key.Keys) { [void](M (G $filter $field) 'Установить' @(, $key[$field])) }
    , $set
}

$source = Connect-Base $SourceRef
$target = Connect-Base $TargetRef

Invoke-Test 'Reception enabled on target' {
    [void](M (G $target 'WS_ОкружениеСервер') 'ВключитьПриём')
    Add-Result 'Reception enabled on target' (M (G $target 'WS_ОкружениеСервер') 'ПриёмВключён')
}

Invoke-Test 'Reception repairs a missing Users item' {
    $environment = G $target 'WS_ОкружениеСервер'
    if ($null -eq (M (G $target 'Метаданные') 'НайтиПоПолномуИмени' @('Справочник.Пользователи'))) {
        Add-Result 'Reception repairs a missing Users item' $true 'no Users catalog'
        return
    }
    $gateUser = M (G $target 'ПользователиИнформационнойБазы') 'НайтиПоИмени' @([string](M $environment 'ИмяСлужебногоПользователя'))
    $item = M (G (G $target 'Справочники') 'Пользователи') 'НайтиПоРеквизиту' @('ИдентификаторПользователяИБ', (G $gateUser 'УникальныйИдентификатор'))
    if (-not (M $item 'Пустая')) {
        $object = M $item 'ПолучитьОбъект'
        S (G $object 'ОбменДанными') 'Загрузка' $true
        [void](M $object 'Удалить')
    }
    $broken = -not (M $environment 'ПриёмНастроенПолностью')
    [void](M $environment 'ВключитьПриём')
    $repaired = [bool](M $environment 'ПриёмНастроенПолностью')
    Add-Result 'Reception repairs a missing Users item' ($broken -and $repaired) "detected=$broken repaired=$repaired"
}

Invoke-Test 'Reception repairs a login disabled by BSP' {
    $environment = G $target 'WS_ОкружениеСервер'
    $gateUser = M (G $target 'ПользователиИнформационнойБазы') 'НайтиПоИмени' @([string](M $environment 'ИмяСлужебногоПользователя'))
    # Так БСП выключает вход неактивному пользователю: снимает стандартную аутентификацию.
    S $gateUser 'АутентификацияСтандартная' $false
    [void](M $gateUser 'Записать')
    $broken = [string](M $environment 'ПроблемаПриёма')
    [void](M $environment 'ВключитьПриём')
    $gateUser = M (G $target 'ПользователиИнформационнойБазы') 'НайтиПоИмени' @([string](M $environment 'ИмяСлужебногоПользователя'))
    $loginBack = [bool](G $gateUser 'АутентификацияСтандартная')
    $after = [string](M $environment 'ПроблемаПриёма')
    Add-Result 'Reception repairs a login disabled by BSP' ($broken -ne '' -and $loginBack -and $after -eq '') "detected='$broken' login=$loginBack after='$after'"
}

Invoke-Test 'Gate password is random and rotates' {
    $targetEnvironment = G $target 'WS_ОкружениеСервер'
    $sourceEnvironment = G $source 'WS_ОкружениеСервер'
    $first = [string](M $targetEnvironment 'ПарольСлужебногоПользователя')
    [void](M $targetEnvironment 'СоздатьНовыйПарольСлужебногоВхода')
    $second = [string](M $targetEnvironment 'ПарольСлужебногоПользователя')
    $changed = $first -ne $second -and $second.Length -ge 60

    # Источник со старым паролем не должен войти: приёмник отклоняет служебный вход.
    # Пароль каждый раз новый: база может запрещать недавно использованные.
    [void](M $sourceEnvironment 'УстановитьПарольСлужебногоВхода' @("Stale-$([guid]::NewGuid().ToString('N'))-1"))
    $refused = M (G $source 'WS_ТранспортСервер') 'СохранитьАдресРабочейБазы' @($TargetUrl)
    $refusedByGate = [string]$refused -match 'служебный вход'

    [void](M $sourceEnvironment 'УстановитьПарольСлужебногоВхода' @($second))
    # Повторное сохранение того же пароля не должно упасть на запрете недавних паролей.
    [void](M $sourceEnvironment 'УстановитьПарольСлужебногоВхода' @($second))
    $accepted = [string](M (G $source 'WS_ТранспортСервер') 'СохранитьАдресРабочейБазы' @($TargetUrl))
    $passedGate = $accepted -notmatch 'служебный вход' -and $accepted -match 'не рабочая база'
    Add-Result 'Gate password is random and rotates' ($changed -and $refusedByGate -and $passedGate) "changed=$changed refused=$refusedByGate accepted=$passedGate"
}

Invoke-Test 'Base name detection' {
    $environment = G $source 'WS_ОкружениеСервер'
    $wrong = @()
    foreach ($case in @(@('base_dev', $true), @('copy_base', $true), @('sales_dev_ivanov', $true), @('test', $true), @('dev01', $true), @('база-копия', $true),
            @('devices', $false), @('kopeyka', $false), @('contest_base', $false), @('sales_main', $false))) {
        if ([bool](M $environment 'ЭтоИмяКопии' @($case[0])) -ne $case[1]) { $wrong += $case[0] }
    }
    Add-Result 'Base name detection' ($wrong.Count -eq 0) "wrong=$($wrong -join ',')"
}

Invoke-Test 'Extension version is reported' {
    $sourceVersion = [string](M (G $source 'WS_ОкружениеСервер') 'ВерсияРасширения')
    $targetVersion = [string](M (G $target 'WS_ОкружениеСервер') 'ВерсияРасширения')
    Add-Result 'Extension version is reported' ($sourceVersion -match '^\d+\.\d+\.\d+$' -and $sourceVersion -eq $targetVersion) "source=$sourceVersion target=$targetVersion"
}

Invoke-Test 'Production address rule' {
    $environment = G $source 'WS_ОкружениеСервер'
    $rule = { param($thisType, $targetType, $targetId) [string](M $environment 'ОшибкаАдресаРабочейБазы' @($thisType, 'this', $targetType, $targetId)) }
    $ok = (& $rule 'production' 'production' 'this') -eq '' -and (& $rule 'production' 'production' 'other') -ne '' -and
        (& $rule 'test' 'production' 'prod') -eq '' -and (& $rule 'test' 'test' 'other') -ne ''
    $rejected = [string](M (G $source 'WS_ТранспортСервер') 'СохранитьАдресРабочейБазы' @($TargetUrl))
    Add-Result 'Production address rule' ($ok -and $rejected -ne '') "rule=$ok rejectedCopy=$rejected"
}

Invoke-Test 'Saved production address is offered' {
    $environment = G $source 'WS_ОкружениеСервер'
    $transport = G $source 'WS_ТранспортСервер'
    $sample = 'http://hkt-test.invalid/base'
    [void](M $environment 'ЗаписатьАдресРабочейБазы' @($sample))
    try {
        $stored = [string](M $environment 'АдресРабочейБазы')
        $offered = @(Get-Items (M $transport 'АдресаДляВыбора')) -contains $sample
    } finally {
        [void](M $transport 'СохранитьАдресРабочейБазы' @(''))
    }
    $cleared = [string](M $environment 'АдресРабочейБазы')
    Add-Result 'Saved production address is offered' ($stored -eq $sample -and $offered -and $cleared -eq '') "stored=$stored offered=$offered cleared='$cleared'"
}

Invoke-Test 'Address list is bound to the base' {
    $sourceKey = [string](M (G $source 'WS_ТранспортСервер') 'КлючАдресов')
    $targetKey = [string](M (G $target 'WS_ТранспортСервер') 'КлючАдресов')
    Add-Result 'Address list is bound to the base' ($sourceKey -ne $targetKey -and $sourceKey -match ([regex]::Escape($SourceRef) + '$')) "source=$sourceKey target=$targetKey"
}

Invoke-Test 'Full list is shown before transfer' {
    $script:linked = New-LinkedCatalogItems $source
    $selection = New-Selection $source @($linked.Child)
    $check = Invoke-Check $source $selection
    $rows = Get-Rows $check
    $parentRow = $rows | Where-Object { G $_ 'Related' } | Select-Object -First 1
    $script:parentKey = [string](G $parentRow 'Key')
    $edges = M (G $parentRow 'Parents') 'Количество'
    Add-Result 'Full list is shown before transfer' ((G $check 'Успешно') -and $rows.Count -eq 2 -and (G $parentRow 'Action') -eq 'create' -and $edges -eq 1) "catalog=$($linked.Catalog) rows=$($rows.Count) edges=$edges err=$(G $check 'Ошибка')"

    $serialization = G $source 'WS_СериализацияСервер'
    $opened = M $serialization 'ЗначениеПоКлючу' @($parentKey, $selection)
    $missingKey = ($parentKey -split '/')[0] + '/' + [guid]::NewGuid().ToString()
    $absent = M $serialization 'ЗначениеПоКлючу' @($missingKey, $selection)
    $sameObject = $null -ne $opened -and (Get-Uuid $source $opened) -eq (Get-Uuid $source $linked.Parent)
    Add-Result 'Row opens its object' ($sameObject -and $null -eq $absent) "same=$sameObject absentIsEmpty=$($null -eq $absent)"
}

Invoke-Test 'Only checked objects are written' {
    $selection = New-Selection $source @($linked.Child)
    $unchecked = New-Selection $source @($parentKey)
    $transfer = Invoke-Transfer $source $selection '' $unchecked
    $child = Test-Exists (Get-RefByUuid $target 'Справочники' $linked.Catalog (Get-Uuid $source $linked.Child))
    $parent = Test-Exists (Get-RefByUuid $target 'Справочники' $linked.Catalog (Get-Uuid $source $linked.Parent))
    Add-Result 'Only checked objects are written' ((G $transfer 'Записано') -and $child -and -not $parent) "child=$child parent=$parent err=$(G $transfer 'Ошибка')"
}

Invoke-Test 'Related object written when checked' {
    $selection = New-Selection $source @($linked.Child)
    $transfer = Invoke-Transfer $source $selection
    $parent = Test-Exists (Get-RefByUuid $target 'Справочники' $linked.Catalog (Get-Uuid $source $linked.Parent))
    Add-Result 'Related object written when checked' ((G $transfer 'Записано') -and $parent) "err=$(G $transfer 'Ошибка')"
}

Invoke-Test 'Repeated selection gives one row' {
    $rows = Get-Rows (Invoke-Check $source (New-Selection $source @($linked.Child, $linked.Child)))
    Add-Result 'Repeated selection gives one row' ($rows.Count -eq 1) "rows=$($rows.Count)"
}

Invoke-Test 'Object versions give their objects once' {
    if ($null -eq (M (G $source 'Метаданные') 'НайтиПоПолномуИмени' @('РегистрСведений.ВерсииОбъектов'))) {
        Add-Result 'Object versions give their objects once' $true 'no ВерсииОбъектов register'
        return
    }
    $text = 'ВЫБРАТЬ ПЕРВЫЕ 2 Версии.Объект КАК Объект, Версии.НомерВерсии КАК НомерВерсии ИЗ РегистрСведений.ВерсииОбъектов КАК Версии ' +
        'ГДЕ Версии.Объект В (ВЫБРАТЬ ПЕРВЫЕ 1 Т.Объект ИЗ РегистрСведений.ВерсииОбъектов КАК Т ' +
        'СГРУППИРОВАТЬ ПО Т.Объект ИМЕЮЩИЕ КОЛИЧЕСТВО(*) > 1)'
    $rowsQuery = M (M (M $source 'NewObject' @('Запрос', $text)) 'Выполнить') 'Выбрать'
    $keys = @()
    while (M $rowsQuery 'Следующий') {
        $keyStructure = M $source 'NewObject' @('Структура')
        [void](M $keyStructure 'Вставить' @('Объект', (G $rowsQuery 'Объект')))
        [void](M $keyStructure 'Вставить' @('НомерВерсии', (G $rowsQuery 'НомерВерсии')))
        $keys += M (G (G $source 'РегистрыСведений') 'ВерсииОбъектов') 'СоздатьКлючЗаписи' @(, $keyStructure)
    }
    if ($keys.Count -lt 2) {
        Add-Result 'Object versions give their objects once' $true 'no object with two versions'
        return
    }
    $selection = New-Selection $source $keys
    $fields = @(Get-Items (M (G $source 'WS_ТранспортСервер') 'ПоляСсылокРегистра' @(, $selection)) | ForEach-Object { G $_ 'Имя' })
    $transport = G $source 'WS_ТранспортСервер'
    $storage = G $source 'ХранилищеОбщихНастроек'
    $fieldList = M $transport 'ПоляСсылокРегистра' @(, $selection)
    [void](M $storage 'Удалить' @('WS_HotKeyTransfer', 'ПолеОбъекта|РегистрСведений.ВерсииОбъектов', $null))
    $default = [string](M $transport 'ПолеОбъектаПоУмолчанию' @($selection, $fieldList))
    [void](M $transport 'ЗапомнитьПолеОбъекта' @($selection, ''))
    $remembered = [string](M $transport 'ПолеОбъектаПоУмолчанию' @($selection, $fieldList))
    [void](M $storage 'Удалить' @('WS_HotKeyTransfer', 'ПолеОбъекта|РегистрСведений.ВерсииОбъектов', $null))
    Add-Result 'Object field is the default and the choice is kept' ($default -eq 'Объект' -and $remembered -eq '') "default='$default' afterRecordsChosen='$remembered'"
    $check = M (G $source 'WS_ТранспортСервер') 'Проверить' @($selection, $TargetUrl, 'Объект')
    $allRows = Get-Rows $check
    $selected = @($allRows | Where-Object { -not (G $_ 'Related') })
    Add-Result 'Object versions give their objects once' (($fields -contains 'Объект') -and $selected.Count -eq 1) "fields=$($fields -join ',') selectedRows=$($selected.Count) err=$(G $check 'Ошибка')"
}

Invoke-Test 'Identical object is unchanged' {
    $rows = Get-Rows (Invoke-Check $source (New-Selection $source @($linked.Child)))
    $action = G $rows[0] 'Action'
    Add-Result 'Identical object is unchanged' ($rows.Count -eq 1 -and $action -eq 'unchanged') "rows=$($rows.Count) action=$action changes=$(G $rows[0] 'Changes')"
}

Invoke-Test 'Address list follows check and transfer' {
    $transport = G $source 'WS_ТранспортСервер'
    $storage = G $source 'ХранилищеОбщихНастроек'
    $key = [string](M $transport 'КлючАдресов')
    $saved = M $transport 'ПолучитьАдреса'
    $sample = 'http://hkt-test.invalid/base'
    $selection = New-Selection $source @($linked.Child)
    try {
        [void](M $storage 'Сохранить' @('WS_HotKeyTransfer', $key, (New-Selection $source @($sample))))
        [void](Invoke-Check $source $selection)
        $afterCheck = @(Get-Items (M $transport 'ПолучитьАдреса')) -join '|'
        [void](M $storage 'Сохранить' @('WS_HotKeyTransfer', $key, (New-Selection $source @('http://first.invalid/base'))))
        [void](M $transport 'Проверить' @($selection, $sample, ''))
        $afterFailure = @(Get-Items (M $transport 'ПолучитьАдреса')) -join '|'
        [void](M $storage 'Сохранить' @('WS_HotKeyTransfer', $key, (New-Selection $source @($sample, $TargetUrl))))
        [void](Invoke-Transfer $source $selection)
        $afterTransfer = @(Get-Items (M $transport 'ПолучитьАдреса')) -join '|'
    } finally {
        [void](M $storage 'Сохранить' @('WS_HotKeyTransfer', $key, $saved))
    }
    $ok = $afterCheck -eq "$sample|$TargetUrl" -and $afterFailure -eq 'http://first.invalid/base' -and $afterTransfer -eq "$TargetUrl|$sample"
    Add-Result 'Address list follows check and transfer' $ok "check=$afterCheck failure=$afterFailure transfer=$afterTransfer"
}

Invoke-Test 'Changed fields are named' {
    $object = M $linked.Child 'ПолучитьОбъект'
    S (G $object 'ОбменДанными') 'Загрузка' $true
    S $object 'Наименование' ("HKT_CHANGED_" + (Get-Date -Format 'HHmmss'))
    [void](M $object 'Записать')
    $rows = Get-Rows (Invoke-Check $source (New-Selection $source @($linked.Child)))
    $changes = [string](G $rows[0] 'Changes')
    Add-Result 'Changed fields are named' ((G $rows[0] 'Action') -eq 'update' -and $changes -eq 'Наименование') "changes=$changes"
}

Invoke-Test 'Posted document restores movements' {
    $found = Find-DocumentWithMovements $source
    $targetRef = Get-RefByUuid $target 'Документы' $found.Document $found.Uuid
    if (Test-Exists $targetRef) {
        $set = G (G (M $targetRef 'ПолучитьОбъект') 'Движения') $found.Register
        [void](M $set 'Прочитать')
        [void](M $set 'Очистить')
        S (G $set 'ОбменДанными') 'Загрузка' $true
        [void](M $set 'Записать')
    }
    $selection = New-Selection $source @($found.Ref)
    $before = Get-Rows (Invoke-Check $source $selection)
    $changes = [string](G $before[0] 'Changes')
    $transfer = Invoke-Transfer $source $selection
    $object = M $targetRef 'ПолучитьОбъект'
    $set = G (G $object 'Движения') $found.Register
    [void](M $set 'Прочитать')
    $after = M $set 'Количество'
    $repeat = Get-Rows (Invoke-Check $source $selection)
    Add-Result 'Posted document restores movements' ((G $transfer 'Записано') -and $after -eq $found.Count -and (G $object 'Проведен') -and $changes -match 'движения') "doc=$($found.Document) register=$($found.Register) expected=$($found.Count) actual=$after changes=$changes err=$(G $transfer 'Ошибка')"
    $script:documentName = $found.Document
    Add-Result 'No false differences after transfer' ((G $repeat[0] 'Action') -eq 'unchanged') "action=$(G $repeat[0] 'Action') changes=$(G $repeat[0] 'Changes')"
}

Invoke-Test 'Transfer is recorded in the event log' {
    $events = M $target 'NewObject' @('ТаблицаЗначений')
    $filter = M $target 'NewObject' @('Структура')
    [void](M $filter 'Вставить' @('СобытиеНачала', (Get-Date).AddMinutes(-15)))
    $eventNames = M $target 'NewObject' @('Массив')
    [void](M $eventNames 'Добавить' @('WS_HotKeyTransfer.Receive'))
    [void](M $filter 'Вставить' @('Событие', $eventNames))
    [void](M $target 'ВыгрузитьЖурналРегистрации' @($events, $filter))
    $found = $false
    foreach ($row in @(Get-Items $events)) {
        $comment = [string](G $row 'Комментарий')
        if ($comment.Contains("Пользователь: $User") -and $comment.Contains($documentName)) { $found = $true }
    }
    Add-Result 'Transfer is recorded in the event log' $found "rows=$(M $events 'Количество')"
}

Invoke-Test 'Register record restored by key' {
    $found = Find-RegisterRecord $source
    $targetKey = @{}
    foreach ($field in $found.Key.Keys) {
        $value = $found.Key[$field]
        if ($value -is [System.__ComObject]) {
            $parts = ([string](M (M $value 'Метаданные') 'ПолноеИмя')).Split('.')
            $collection = @{ 'Справочник' = 'Справочники'; 'Документ' = 'Документы' }[$parts[0]]
            $targetKey[$field] = Get-RefByUuid $target $collection $parts[1] (Get-Uuid $source $value)
        } else { $targetKey[$field] = $value }
    }
    $set = New-Filtered-Set $target $found.Register $targetKey
    S (G $set 'ОбменДанными') 'Загрузка' $true
    [void](M $set 'Записать')
    $keyStructure = M $source 'NewObject' @('Структура')
    foreach ($field in $found.Key.Keys) { [void](M $keyStructure 'Вставить' @($field, $found.Key[$field])) }
    $recordKey = M (G (G $source 'РегистрыСведений') $found.Register) 'СоздатьКлючЗаписи' @(, $keyStructure)
    $recordSelection = New-Selection $source @($recordKey)
    $recordRows = Get-Rows (Invoke-Check $source $recordSelection)
    $openedRecord = M (G $source 'WS_СериализацияСервер') 'ЗначениеПоКлючу' @([string](G $recordRows[0] 'Key'), $recordSelection)
    Add-Result 'Register row opens its record' ($null -ne $openedRecord) "rows=$($recordRows.Count)"
    $transfer = Invoke-Transfer $source $recordSelection
    $sourceSet = New-Filtered-Set $source $found.Register $found.Key
    [void](M $sourceSet 'Прочитать')
    $targetSet = New-Filtered-Set $target $found.Register $targetKey
    [void](M $targetSet 'Прочитать')
    $expected = M $sourceSet 'Количество'
    $actual = M $targetSet 'Количество'
    Add-Result 'Register record restored by key' ((G $transfer 'Записано') -and $actual -eq $expected -and $actual -gt 0) "register=$($found.Register) expected=$expected actual=$actual err=$(G $transfer 'Ошибка')"
}

Invoke-Test 'Same base is blocked' {
    [void](M (G $source 'WS_ОкружениеСервер') 'ВключитьПриём')
    $selection = New-Selection $source @($linked.Child)
    $sameUrl = $TargetUrl.Replace($TargetRef, $SourceRef)
    $check = M (G $source 'WS_ТранспортСервер') 'Проверить' @($selection, $sameUrl, '')
    $code = G $check 'КодОшибки'
    Add-Result 'Same base is blocked' ($code -eq 'WS_SAME_ENVIRONMENT') "code=$code err=$(G $check 'Ошибка')"
}

function New-ForgedRequest($v, [string] $signature, [string] $time, [string] $targetId, [string] $userName) {
    $id = if ($targetId) { $targetId } else { [string](M (G $v 'WS_ОкружениеСервер') 'ПолучитьИдентификатор') }
    @{ ProtocolVersion = 3; User = $userName; Source = 'forged'; Target = $id; Time = $time; Reason = ''; Signature = $signature; Body = '{"Items":[],"AllowedMissing":[]}' } | ConvertTo-Json -Compress
}

$now = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

Invoke-Test 'Forged signature is rejected' {
    $request = New-ForgedRequest $target 'forged-signature' $now '' $User
    $answer = M (G $target 'WS_СериализацияСервер') 'СообщениеИзJSON' @([string](M (G $target 'WS_ПриёмникСервер') 'ПрименитьПакет' @($request)))
    $code = G $answer 'Code'
    Add-Result 'Forged signature is rejected' ($code -eq 'WS_ACCESS_DENIED') "code=$code"
}

Invoke-Test 'Unknown user is rejected' {
    $request = New-ForgedRequest $target 'forged-signature' $now '' 'no_such_user_hkt'
    $answer = M (G $target 'WS_СериализацияСервер') 'СообщениеИзJSON' @([string](M (G $target 'WS_ПриёмникСервер') 'ПрименитьПакет' @($request)))
    $code = G $answer 'Code'
    Add-Result 'Unknown user is rejected' ($code -eq 'WS_ACCESS_DENIED') "code=$code"
}

Invoke-Test 'Expired request is rejected' {
    $request = New-ForgedRequest $target 'forged-signature' '2000-01-01T00:00:00' '' $User
    $answer = M (G $target 'WS_СериализацияСервер') 'СообщениеИзJSON' @([string](M (G $target 'WS_ПриёмникСервер') 'ПрименитьПакет' @($request)))
    $code = G $answer 'Code'
    Add-Result 'Expired request is rejected' ($code -eq 'WS_REQUEST_EXPIRED') "code=$code"
}

Invoke-Test 'Request for another base is rejected' {
    $request = New-ForgedRequest $target 'forged-signature' $now 'another/base' $User
    $answer = M (G $target 'WS_СериализацияСервер') 'СообщениеИзJSON' @([string](M (G $target 'WS_ПриёмникСервер') 'ПрименитьПакет' @($request)))
    $code = G $answer 'Code'
    Add-Result 'Request for another base is rejected' ($code -eq 'WS_TARGET_MISMATCH') "code=$code"
}

Invoke-Test 'Service account cannot open the base' {
    $gateUser = [string](M (G $target 'WS_ОкружениеСервер') 'ИмяСлужебногоПользователя')
    $gatePassword = [string](M (G $target 'WS_ОкружениеСервер') 'ПарольСлужебногоПользователя')
    try {
        $gate = $connector.Connect("Srvr=`"$Server`";Ref=`"$TargetRef`";Usr=`"$gateUser`";Pwd=`"$gatePassword`"")
        Add-Result 'Service account cannot open the base' $false 'external connection was allowed'
    } catch {
        Add-Result 'Service account cannot open the base' $true 'external connection refused'
    }
}

$results | Format-Table -AutoSize -Wrap
if ($results.Passed -contains $false) { exit 1 }
