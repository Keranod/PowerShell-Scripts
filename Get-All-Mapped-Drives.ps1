$usersSID = Get-ChildItem "REGISTRY::HKEY_USERS" -Name -Include "S-1-5-21*" -Exclude "*Classes"

foreach ($user in $usersSID) {
    $user
    Get-ChildItem "REGISTRY::HKEY_USERS\$user\Network"
}