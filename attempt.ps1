# Import the Active Directory module
Import-Module ActiveDirectory

# Define an empty array to store user data
$usersData = @()

# Step 1: Query Active Directory for a full list of AD users
$users = Get-ADUser -Filter * -Properties SamAccountName, Name, Surname, Description

# Step 2: Iterate through each user to get the groups they belong to
foreach ($user in $users) {
    $userGroups = Get-ADPrincipalGroupMembership -Identity $user | Select-Object -ExpandProperty Name
    $userGroupsList = $userGroups -join ", "
    
    # Create an object containing user data and their groups
    $userObject = New-Object PSObject -Property @{
        SamAccountName = $user.SamAccountName
        Name = $user.Name
        Surname = $user.Surname
        Description = $user.Description
        Groups = $userGroupsList
    }
    
    # Add the user object to the array
    $usersData += $userObject
}

# Step 3: Export the data to a .csv file
$usersData | Export-Csv -Path "AD_User_Groups.csv" -NoTypeInformation
