$SourcePath = 'C:\Sources'
$Sources = @('Applications'; 'Packages'; 'Operating System Images'; 'Operating System Upgrades'; 'Software Updates'; 'Drivers'; 'Security Roles'; 'Export-Import' )

ForEach($Source in $Sources){
    New-Item -Name $Source -ItemType Directory -Path "$SourcePath"
}