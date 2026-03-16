param(
  [Parameter(Mandatory = $true)]
  [string] $WebUrl,

  [Parameter(Mandatory = $false)]
  [string] $ListUrl = 'Lists/EditorLinks',

  [Parameter(Mandatory = $false)]
  [string] $FieldInternalName = 'PzlOfficeUIFabricIcon',

  [Parameter(Mandatory = $false)]
  [string] $ComponentId = '8b14c246-5920-4cf2-bc89-2514d59b74f9'
)

$ErrorActionPreference = 'Stop'

Write-Host "Associating field customizer with '$FieldInternalName' on '$ListUrl' in '$WebUrl'..."
Write-Host 'Prerequisite: run `m365 login` before executing this script.'
Write-Host 'Note: this script sets the list-view field customizer only. It does not hide the field in the standard New/Edit/Display forms.'

$fieldJson = m365 spo field get --webUrl $WebUrl --listUrl $ListUrl --internalName $FieldInternalName --output json

if (-not $fieldJson) {
  throw "Could not retrieve field '$FieldInternalName' from list '$ListUrl'."
}

$field = $fieldJson | ConvertFrom-Json

if (-not $field.InternalName) {
  throw "The field lookup did not return a valid field payload for '$FieldInternalName'."
}

m365 spo field set --webUrl $WebUrl --listUrl $ListUrl --internalName $FieldInternalName --ClientSideComponentId $ComponentId

$updatedFieldJson = m365 spo field get --webUrl $WebUrl --listUrl $ListUrl --internalName $FieldInternalName --output json
$updatedField = $updatedFieldJson | ConvertFrom-Json

Write-Host ''
Write-Host 'Updated field:'
Write-Host "  List URL: $ListUrl"
Write-Host "  Internal name: $($updatedField.InternalName)"
Write-Host "  ClientSideComponentId: $($updatedField.ClientSideComponentId)"
