#http://www.mibdepot.com/cgi-bin/getmib3.cgi?win=mib_a&i=1759&n=Printer-MIB&r=cisco&f=Printer-MIB-V1SMI.my&v=v1&t=tree#prtMarkerSuppliesGroup
enum SupplyType {
    Other = 1
    Unknown = 2
    Toner = 3
    WasteToner = 4
    Ink = 5
    InkCartridge = 6
    ImagingDrum = 9
    Developer = 10
    Fuser = 15
    CleanerUnit = 18
    TransferUnit = 20
    TonerCartridge = 21
}
 
enum SupplyUnit {
    Impressions = 7
    TenthsOfGrams = 13
    TenthsOfMilliliters = 15
    Percent = 19
}
 
enum Severity {
    Other = 1
    Critical = 3
    Warn = 4
    Warning = 5
}
 
enum Training {
    Other = 1
    Unknown = 2
    Untrained = 3
    Trained = 4
    FieldService = 5
    Management = 6
    NoInterventionRequired = 7
}
 
enum Code {
    Other = 1
    Unknown = 2
    CoverOpen = 3
    CoverClosed = 4
    InterlockOpen = 5
    InterlockClosed = 6
    ConfigurationChange = 7
    Jam = 8
    SubunitMissing = 9
    SubunitLifeAlmostOver = 10
    SubunitLifeOver = 11
    SubunitAlmostEmpty = 12
    SubunitEmpty = 13
    SubunitAlmostFull = 14
    SubunitFull = 15
    SubunitNearLimit = 16
    SubunitAtLimit = 17
    SubunitOpened = 18
    SubunitClosed = 19
    SubunitTurnedOn = 20
    SubunitTurnedOff = 21
    SubunitOffline = 22
    SubunitPowerSaver = 23
    SubunitWarmingUp = 24
    SubunitAdded = 25
    SubunitRemoved = 26
    SubunitResourceAdded = 27
    SubunitResourceRemoved = 28
    SubunitRecoverableFailure = 29
    SubunitUnrecoverableFailure = 30
    SubunitRecoverableStorageError = 31
    SubunitUnrecoverableStorageError = 32
    SubunitMotorFailure = 33
    SubunitMemoryExhausted = 34
    SubunitUnderTemperature = 35
    SubunitOverTemperature = 36
    SubunitTimingFailure = 37
    SubunitThermistorFailure = 38
    DoorOpen = 501
    DoorClosed = 502
    PowerUp = 503
    PowerDown = 504
    PrinterNMSReset = 505
    PrinterManualReset = 506
    PrinterReadyToPrint = 507
    InputMediaTrayMissing = 801
    InputMediaSizeChange = 802
    InputMediaWeightChange = 803
    InputMediaTypeChange = 804
    InputMediaColorChange = 805
    InputMediaFormPartsChange = 806
    InputMediaSupplyLow = 807
    InputMediaSupplyEmpty = 808
    InputMediaChangeRequest = 809
    InputManualInputRequest = 810
    InputTrayPositionFailure = 811
    InputTrayElevationFailure = 812
    InputCannotFeedSizeSelected = 813
    OutputMediaTrayMissing = 901
    OutputMediaTrayAlmostFull = 902
    OutputMediaTrayFull = 903
    OutputMailboxSelectFailure = 904
    MarkerFuserUnderTemperature = 1001
    MarkerFuserOverTemperature = 1002
    MarkerFuserTimingFailure = 1003
    MarkerFuserThermistorFailure = 1004
    MarkerAdjustingPrintQuality = 1005
    MarkerTonerEmpty = 1101
    MarkerInkEmpty = 1102
    MarkerPrintRibbonEmpty = 1103
    MarkerTonerAlmostEmpty = 1104
    MarkerInkAlmostEmpty = 1105
    MarkerPrintRibbonAlmostEmpty = 1106
    MarkerWasteTonerReceptacleAlmostFull = 1107
    MarkerWasteInkReceptacleAlmostFull = 1108
    MarkerWasteTonerReceptacleFull = 1109
    MarkerWasteInkReceptacleFull = 1110
    MarkerOpcLifeAlmostOver = 1111
    MarkerOpcLifeOver = 1112
    MarkerDeveloperAlmostEmpty = 1113
    MarkerDeveloperEmpty = 1114
    MarkerTonerCartridgeMissing = 1115
    MediaPathMediaTrayMissing = 1301
    MediaPathMediaTrayAlmostFull = 1302
    MediaPathMediaTrayFull = 1303
    MediaPathCannotDuplexMediaSelected = 1304
    InterpreterMemoryIncrease = 1501
    InterpreterMemoryDecrease = 1502
    InterpreterCartridgeAdded = 1503
    InterpreterCartridgeDeleted = 1504
    InterpreterResourceAdded = 1505
    InterpreterResourceDeleted = 1506
    InterpreterResourceUnavailable = 1507
    InterpreterComplexPageEncountered = 1509
    AlertRemovalOfBinaryChangeEntry = 1801
}
 
function Query-Printer {
    param ([string]$IP)
    if (Test-Connection -ComputerName $IP -Count 1) {
        $date = (Get-Date).ToString("yyyy-MM-dd")
        $dateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
        $result = [PSCustomObject]@{
            Model = $null
            Serial = $null
            Alerts = @()
            Faults = @()
            Impressions = @()
            Supplies = @()
        }
        $snmp = New-Object -ComObject olePrn.OleSNMP
        $snmp.open($IP,'public',2,2000)
 
        $result.Model = $snmp.Get('.1.3.6.1.4.1.1602.4.7')
 
        $suppliesTypes = $snmp.GetTree('43.11.1.1.5')
        $suppliesDescriptions = $snmp.GetTree('43.11.1.1.6')
        $suppliesUnits = $snmp.GetTree('43.11.1.1.7')
        $suppliesCapacities = $snmp.GetTree('43.11.1.1.8')
        $suppliesLevels = $snmp.GetTree('43.11.1.1.9')
 
        $alertSeverity = $snmp.GetTree('43.18.1.1.2')
        $alertTraining = $snmp.GetTree('43.18.1.1.3')
        $alertCode = $snmp.GetTree('43.18.1.1.7')
        $alertDescription = $snmp.GetTree('43.18.1.1.8')
        
        $faultDescription = $snmp.GetTree('.1.3.6.1.4.1.253.8.53.8.2.1.4')
 
        if ($result.Model -like '*Xerox*') {
            $result.Serial = $snmp.Get('.1.3.6.1.4.1.253.8.53.3.2.1.3.1')
            $totalImpressions = $snmp.Get('.1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.1')
            $blackImpressions = $snmp.Get('.1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.34')
            $colorImpressions = $snmp.Get('.1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.33')
            $largeImpressions = $snmp.Get('.1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.47')
            $blackLargeImpressions = $snmp.Get('.1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.44')
            $colorLargeImpressions = $snmp.Get('.1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.43')
        }
        else {
            $result.Serial = $snmp.Get('.1.3.6.1.2.1.43.5.1.1.17.1')
            $totalImpressions = $snmp.Get('.1.3.6.1.2.1.43.10.2.1.4.1.1')
        }
 
        $snmp.close()
 
        $length = $suppliesTypes.GetLength(1)
        for ($i = 0;$i -lt $length; $i++) {
            $capacity = $suppliesCapacities[1,$i]
            if ($capacity -eq '-1') {$capacity = 'Other'}
            if ($capacity -eq '-2') {$capacity = 'Unknown'}
            $level = $suppliesLevels[1,$i]
            if ($level -eq '-2') {$level = 'Unknown'}
            if ($level -eq '-3') {$level = 'BelowCapacity'}
 
            $result.Supplies += [PSCustomObject]@{
                Date = $date
                DateTime = $dateTime
                IP = $IP
                Type = ([SupplyType]($suppliesTypes[1,$i])).ToString()
                Description = (($suppliesDescriptions[1,$i]) -split ';')[0]
                Level = $level.ToString()
                Capacity = $capacity.ToString()        
                Unit = ([SupplyUnit]($suppliesUnits[1,$i])).ToString()
            }
        }
 
        $length = $alertSeverity.GetLength(1)
        for ($i = 0;$i -lt $length; $i++) {
            $training = ([Training]$alertTraining[1,$i]).ToString()
        }
    }
}