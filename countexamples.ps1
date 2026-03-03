$categories = @("hssf","xssf","ss","xwpf","poifs","ooxml","sxssf","scratchpad")

$all = foreach ($c in $categories) {
  if (Test-Path $c) {
    Get-ChildItem -Path $c -Directory | ForEach-Object {
      [PSCustomObject]@{
        Category = $c
        Name     = $_.Name
        Path     = $_.FullName
      }
    }
  }
}

"Total example folders: $($all.Count)"
$all | Group-Object Category | Select-Object Name,Count | Format-Table -AutoSize