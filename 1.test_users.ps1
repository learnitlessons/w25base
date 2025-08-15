# COMPLETE LAB ENVIRONMENT SETUP SCRIPT FOR WINDOWS SERVER 2025 GROUPS LESSON
# Author: Lab environment for studying group types
# Domain: Auto-detected
# Requirements: Run on domain controller with administrator rights

Write-Host "🚀 STARTING COMPLETE LAB ENVIRONMENT SETUP FOR GROUPS LESSON" -ForegroundColor Magenta
Write-Host "=================================================================" -ForegroundColor Magenta
Write-Host "This script will create a complete lab environment to demonstrate" -ForegroundColor White
Write-Host "differences between security groups and distribution groups" -ForegroundColor White
Write-Host "=================================================================" -ForegroundColor Magenta

# Get domain information
$domain = Get-ADDomain
$domainDN = $domain.DistinguishedName
Write-Host "`n✅ Domain detected: $($domain.DNSRoot)" -ForegroundColor Green
Write-Host "✅ Distinguished Name: $domainDN" -ForegroundColor Green

# STAGE 1: CREATE ORGANIZATIONAL UNITS
Write-Host "`n🔍 STAGE 1: Creating organizational unit structure..." -ForegroundColor Cyan
Write-Host "Creating OUs for lab work and departments" -ForegroundColor Gray

# Create main OU for lab work
try {
    $labGroupsOU = Get-ADOrganizationalUnit -Identity "OU=Lab-Groups,$domainDN" -ErrorAction Stop
    Write-Host "⚠️ OU Lab-Groups already exists: $($labGroupsOU.DistinguishedName)" -ForegroundColor Yellow
} catch {
    try {
        New-ADOrganizationalUnit -Name "Lab-Groups" -Path $domainDN -Description "Lab work for Windows Server 2025 groups"
        Write-Host "✅ OU Lab-Groups successfully created" -ForegroundColor Green
    } catch {
        Write-Host "❌ CRITICAL ERROR: Failed to create Lab-Groups: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Check access rights and try again" -ForegroundColor Red
        return
    }
}

# Create department sub-units
$departments = @(
    @{Name="IT"; Description="Information Technology Department"},
    @{Name="Finance"; Description="Finance Department"},
    @{Name="Marketing"; Description="Marketing Department"},
    @{Name="HR"; Description="Human Resources Department"}
)

foreach ($dept in $departments) {
    try {
        $deptOU = Get-ADOrganizationalUnit -Identity "OU=$($dept.Name),OU=Lab-Groups,$domainDN" -ErrorAction Stop
        Write-Host "⚠️ OU $($dept.Name) already exists" -ForegroundColor Yellow
    } catch {
        try {
            New-ADOrganizationalUnit -Name $dept.Name -Path "OU=Lab-Groups,$domainDN" -Description $dept.Description
            Write-Host "✅ OU $($dept.Name) created" -ForegroundColor Green
        } catch {
            Write-Host "❌ Error creating OU $($dept.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Display created OU structure
Write-Host "`n📋 Created organizational unit structure:" -ForegroundColor Cyan
try {
    $allOUs = Get-ADOrganizationalUnit -Filter "Name -like '*'" -SearchBase "OU=Lab-Groups,$domainDN" | Select-Object Name, DistinguishedName | Sort-Object Name
    $allOUs | Format-Table -AutoSize
    Write-Host "✅ STAGE 1 COMPLETED: OU structure ready ($($allOUs.Count) units)" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to display OU structure" -ForegroundColor Red
}

# STAGE 2: CREATE TEST USERS
Write-Host "`n👥 STAGE 2: Creating test users..." -ForegroundColor Cyan
Write-Host "Creating users for group demonstration" -ForegroundColor Gray

# Check OU readiness before creating users
$departmentNames = @("IT", "Finance", "Marketing", "HR")
$allOUReady = $true
foreach ($deptName in $departmentNames) {
    try {
        Get-ADOrganizationalUnit -Identity "OU=$deptName,OU=Lab-Groups,$domainDN" -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "❌ OU $deptName not ready!" -ForegroundColor Red
        $allOUReady = $false
    }
}

if (-not $allOUReady) {
    Write-Host "❌ CRITICAL ERROR: Not all OUs created. Stopping execution!" -ForegroundColor Red
    return
}

# Set up passwords and create users
$userPassword = ConvertTo-SecureString "Somepass1" -AsPlainText -Force
Write-Host "🔑 Password for all users: Somepass1" -ForegroundColor Yellow

# Define users to create (extended list)
$usersToCreate = @(
    # IT department - 4 people
    @{Name="JohnSmith"; GivenName="John"; Surname="Smith"; Department="IT"; Title="System Administrator"},
    @{Name="MaryJohnson"; GivenName="Mary"; Surname="Johnson"; Department="IT"; Title="Developer"},
    @{Name="MichaelBrown"; GivenName="Michael"; Surname="Brown"; Department="IT"; Title="Network Administrator"},
    @{Name="SarahDavis"; GivenName="Sarah"; Surname="Davis"; Department="IT"; Title="QA Tester"},
    
    # Finance department - 4 people
    @{Name="RobertWilson"; GivenName="Robert"; Surname="Wilson"; Department="Finance"; Title="Chief Accountant"},
    @{Name="LisaMiller"; GivenName="Lisa"; Surname="Miller"; Department="Finance"; Title="Financial Analyst"},
    @{Name="DavidMoore"; GivenName="David"; Surname="Moore"; Department="Finance"; Title="Economist"},
    @{Name="JenniferTaylor"; GivenName="Jennifer"; Surname="Taylor"; Department="Finance"; Title="Auditor"},
    
    # Marketing department - 3 people
    @{Name="ChrisAnderson"; GivenName="Chris"; Surname="Anderson"; Department="Marketing"; Title="Marketing Specialist"},
    @{Name="AmandaThomas"; GivenName="Amanda"; Surname="Thomas"; Department="Marketing"; Title="PR Manager"},
    @{Name="KevinJackson"; GivenName="Kevin"; Surname="Jackson"; Department="Marketing"; Title="SMM Specialist"},
    
    # HR department - 3 people
    @{Name="LauraWhite"; GivenName="Laura"; Surname="White"; Department="HR"; Title="HR Specialist"},
    @{Name="EmilyMartin"; GivenName="Emily"; Surname="Martin"; Department="HR"; Title="Recruiter"},
    @{Name="JamesGarcia"; GivenName="James"; Surname="Garcia"; Department="HR"; Title="Personnel Manager"}
)

$createdCount = 0
$existingCount = 0

Write-Host "📊 Planning to create users:" -ForegroundColor Cyan
$departments.ForEach({ 
    $deptUsers = $usersToCreate | Where-Object { $_.Department -eq $_.Name }
    Write-Host "  $($_.Name): $($deptUsers.Count) users" -ForegroundColor Gray
})
Write-Host "  Total: $($usersToCreate.Count) users" -ForegroundColor White

foreach ($user in $usersToCreate) {
    Write-Host "👤 Processing user: $($user.GivenName) $($user.Surname) ($($user.Department))" -ForegroundColor Cyan
    
    # Check if user exists
    $existingUser = Get-ADUser -Filter "SamAccountName -eq '$($user.Name)'" -ErrorAction SilentlyContinue
    
    if ($existingUser) {
        Write-Host "⚠️ User $($user.GivenName) $($user.Surname) already exists" -ForegroundColor Yellow
        $existingCount++
    } else {
        try {
            $userParameters = @{
                Name = $user.Name
                GivenName = $user.GivenName
                Surname = $user.Surname
                DisplayName = "$($user.GivenName) $($user.Surname)"
                SamAccountName = $user.Name
                UserPrincipalName = "$($user.Name)@$($domain.DNSRoot)"
                EmailAddress = "$($user.Name.ToLower())@$($domain.DNSRoot)"
                Department = $user.Department
                Title = $user.Title
                Company = "Learn IT Lessons"
                Office = "$($user.Department) Department"
                Path = "OU=$($user.Department),OU=Lab-Groups,$domainDN"
                AccountPassword = $userPassword
                Enabled = $true
                PasswordNeverExpires = $true
                CannotChangePassword = $false
                Description = "Test user for groups lab work"
            }
            
            New-ADUser @userParameters
            Write-Host "✅ $($user.GivenName) $($user.Surname) successfully created" -ForegroundColor Green
            $createdCount++
        } catch {
            Write-Host "❌ ERROR creating $($user.GivenName) $($user.Surname): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Check and report on created users
Write-Host "`n📋 User report by departments:" -ForegroundColor Cyan
try {
    $allUsers = Get-ADUser -Filter * -SearchBase "OU=Lab-Groups,$domainDN" -Properties Department, Title, EmailAddress | 
                Select-Object Name, SamAccountName, Department, Title, EmailAddress, DistinguishedName | 
                Sort-Object Department, Name
    
    if ($allUsers) {
        # Group by departments for nice display
        $departmentNames | ForEach-Object {
            $deptName = $_
            $deptUsers = $allUsers | Where-Object { $_.Department -eq $deptName }
            if ($deptUsers) {
                Write-Host "`n🏢 Department $deptName ($($deptUsers.Count) people):" -ForegroundColor Yellow
                $deptUsers | Select-Object Name, SamAccountName, Title, EmailAddress | Format-Table -AutoSize
            }
        }
        
        Write-Host "`n📊 FINAL STATISTICS:" -ForegroundColor Cyan
        Write-Host "✅ Total users in lab environment: $($allUsers.Count)" -ForegroundColor Green
        Write-Host "✅ Created new: $createdCount, Already existed: $existingCount" -ForegroundColor Green
        
        # Statistics by departments
        $departmentNames | ForEach-Object {
            $deptCount = ($allUsers | Where-Object { $_.Department -eq $_ }).Count
            Write-Host "   $_ : $deptCount users" -ForegroundColor Gray
        }
    } else {
        Write-Host "❌ No users found in OU Lab-Groups" -ForegroundColor Red
    }
} catch {
    Write-Host "❌ Error checking users: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "✅ STAGE 2 COMPLETED: Users ready for demonstration" -ForegroundColor Green

# STAGE 3: CREATE SHARED FOLDERS AND FILES
Write-Host "`n📁 STAGE 3: Creating file structure for demonstration..." -ForegroundColor Cyan
Write-Host "Creating folders and files for testing group permissions" -ForegroundColor Gray

# Create main folder for shared resources
$mainFolder = "C:\LabShares"
if (Test-Path $mainFolder) {
    Write-Host "⚠️ Main folder $mainFolder already exists" -ForegroundColor Yellow
} else {
    try {
        New-Item -Path $mainFolder -ItemType Directory -Force | Out-Null
        Write-Host "✅ Main folder $mainFolder created" -ForegroundColor Green
    } catch {
        Write-Host "❌ Error creating main folder: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Create folders for different departments
$sharedFolders = @(
    @{Path="C:\LabShares\IT-Resources"; Description="IT Department Resources"; Content="IT DEPARTMENT CONFIDENTIAL INFORMATION`nAccess only for IT department employees`n`nContains:`n- Automation scripts`n- Technical documentation`n- Administration tools`n- Configuration files`n`nResponsible:`n- John Smith (System Administrator)`n- Michael Brown (Network Administrator)`n`nCreated: $(Get-Date -Format 'dd.MM.yyyy HH:mm')"},
    
    @{Path="C:\LabShares\Finance-Data"; Description="Financial Data"; Content="COMPANY FINANCIAL DATA`n`nBudget for 2025`nExpense reports for current quarter`nRevenue and investment plans`nAnalytical reports`n`nACCESS ONLY FOR FINANCE DEPARTMENT`n`nResponsible:`n- Robert Wilson (Chief Accountant)`n- Lisa Miller (Financial Analyst)`n- David Moore (Auditor)`n`nLast updated: $(Get-Date -Format 'dd.MM.yyyy HH:mm')"},
    
    @{Path="C:\LabShares\Company-Announcements"; Description="Company Announcements"; Content="ANNOUNCEMENTS FOR ALL EMPLOYEES`n`nDear Colleagues!`n`nNext Friday, $(Get-Date -AddDays 7 -Format 'dd.MM.yyyy'), we will have a company meeting.`nTime: 2:00 PM`nLocation: Conference Room (2nd floor)`nTopic: 'New Windows Server 2025 features and IT policy updates'`n`nAgenda:`n1. New Active Directory features presentation`n2. Security policy updates`n3. IT infrastructure modernization plans`n4. Q&A session`n`nPlease confirm attendance by Thursday.`n`nBest regards,`nIT Department and Administration`n`nContacts:`n- John Smith: john.smith@$($domain.DNSRoot)`n- Laura White: laura.white@$($domain.DNSRoot)"},
    
    @{Path="C:\LabShares\Marketing-Materials"; Description="Marketing Materials"; Content="MARKETING MATERIALS AND RESOURCES`n`nPresentations and brand book`nAdvertising material layouts`nSocial media content`nCampaign analytics and reports`n`nResponsible:`n- Chris Anderson (Marketing Specialist)`n- Amanda Thomas (PR Manager)`n- Kevin Jackson (SMM Specialist)`n`nCreated: $(Get-Date -Format 'dd.MM.yyyy HH:mm')"}
)

foreach ($folder in $sharedFolders) {
    if (Test-Path $folder.Path) {
        Write-Host "⚠️ Folder $($folder.Path) already exists" -ForegroundColor Yellow
    } else {
        try {
            New-Item -Path $folder.Path -ItemType Directory -Force | Out-Null
            Write-Host "✅ Folder created: $($folder.Path)" -ForegroundColor Green
        } catch {
            Write-Host "❌ Error creating folder $($folder.Path): $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
    }
    
    # Create test file in folder
    $fileName = Split-Path $folder.Path -Leaf
    $filePath = Join-Path $folder.Path "$fileName.txt"
    
    try {
        $folder.Content | Out-File $filePath -Encoding UTF8 -Force
        Write-Host "  📄 File created: $filePath" -ForegroundColor Gray
    } catch {
        Write-Host "  ❌ Error creating file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Create additional test file for group conversion demonstration
$testFilePath = "C:\LabShares\TestFile.txt"
$testFileContent = @"
TEST FILE FOR GROUP PERMISSIONS DEMONSTRATION
=============================================

This file is used in lab work to demonstrate:
- Assigning permissions to security groups
- Converting between group types
- Access token behavior
- Impact of group type on resource access

Created: $(Get-Date -Format 'dd.MM.yyyy HH:mm')
Lab environment: $($domain.DNSRoot)
Total users: $($usersToCreate.Count)

Demonstration instructions:
1. Assign permissions to security group
2. Test user access from group
3. Convert group to distribution group
4. Check access change (should disappear)
5. Convert back to security group
6. Check access restoration

Test users by departments:
IT: John Smith, Mary Johnson, Michael Brown, Sarah Davis
Finance: Robert Wilson, Lisa Miller, David Moore, Jennifer Taylor
Marketing: Chris Anderson, Amanda Thomas, Kevin Jackson
HR: Laura White, Emily Martin, James Garcia

© Learn IT Lessons - Windows Server 2025
"@

try {
    $testFileContent | Out-File $testFilePath -Encoding UTF8 -Force
    Write-Host "✅ Special test file created: $testFilePath" -ForegroundColor Green
} catch {
    Write-Host "❌ Error creating test file: $($_.Exception.Message)" -ForegroundColor Red
}

# Display created file structure
Write-Host "`n📋 Created file structure:" -ForegroundColor Cyan
try {
    $fileStructure = Get-ChildItem "C:\LabShares" -Recurse | Select-Object Mode, Name, Length, FullName | Sort-Object FullName
    $fileStructure | Format-Table -AutoSize
    
    $folderCount = ($fileStructure | Where-Object {$_.Mode -like "d*"}).Count
    $fileCount = ($fileStructure | Where-Object {$_.Mode -notlike "d*"}).Count
    Write-Host "✅ Created folders: $folderCount, files: $fileCount" -ForegroundColor Green
} catch {
    Write-Host "❌ Error displaying file structure: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "✅ STAGE 3 COMPLETED: File structure ready" -ForegroundColor Green

# STAGE 4: CONFIGURE REMOTE ACCESS
Write-Host "`n🔐 STAGE 4: Configuring remote access..." -ForegroundColor Cyan
Write-Host "Adding users to groups for RDP and PowerShell Remoting" -ForegroundColor Gray

# Get list of created users
try {
    $labUsers = Get-ADUser -Filter * -SearchBase "OU=Lab-Groups,$domainDN" | Select-Object -ExpandProperty SamAccountName
    
    if ($labUsers.Count -eq 0) {
        Write-Host "❌ Test users not found! Skipping remote access configuration." -ForegroundColor Red
    } else {
        Write-Host "👥 Found users for configuration: $($labUsers.Count)" -ForegroundColor Green
        Write-Host "Users by departments:" -ForegroundColor Gray
        
        $departmentNames | ForEach-Object {
            $deptName = $_
            $deptUsers = Get-ADUser -Filter "Department -eq '$deptName'" -SearchBase "OU=Lab-Groups,$domainDN" -Properties Department | Select-Object -ExpandProperty SamAccountName
            if ($deptUsers) {
                Write-Host "  $deptName ($($deptUsers.Count)): $($deptUsers -join ', ')" -ForegroundColor Gray
            }
        }
        
        # Function to safely add users to security group
        function Add-UsersToSecurityGroup {
            param(
                [string]$GroupName,
                [array]$UserList,
                [string]$GroupDescription
            )
            
            Write-Host "`n🔧 Processing group: $GroupName" -ForegroundColor Cyan
            Write-Host "   Purpose: $GroupDescription" -ForegroundColor Gray
            
            try {
                # Check if group exists
                $targetGroup = Get-ADGroup -Identity $GroupName -ErrorAction Stop
                
                $addedUsers = 0
                $existingUsers = 0
                $errorUsers = 0
                
                foreach ($username in $UserList) {
                    try {
                        # Check if user is already a member
                        $currentMembers = Get-ADGroupMember -Identity $GroupName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SamAccountName
                        
                        if ($currentMembers -contains $username) {
                            Write-Host "  ⚠️ $username is already a group member" -ForegroundColor Yellow
                            $existingUsers++
                        } else {
                            Add-ADGroupMember -Identity $GroupName -Members $username -ErrorAction Stop
                            Write-Host "  ✅ $username added to group" -ForegroundColor Green
                            $addedUsers++
                        }
                    } catch {
                        Write-Host "  ❌ Error adding $username : $($_.Exception.Message)" -ForegroundColor Red
                        $errorUsers++
                    }
                }
                
                Write-Host "  📊 Summary: added $addedUsers, already existed $existingUsers, errors $errorUsers" -ForegroundColor Cyan
                
            } catch {
                Write-Host "  ❌ Group $GroupName not found: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "  💡 This may be normal for some Windows Server configurations" -ForegroundColor Yellow
            }
        }
        
        # Add to remote access groups
        Add-UsersToSecurityGroup -GroupName "Remote Desktop Users" -UserList $labUsers -GroupDescription "Allows Remote Desktop Protocol (RDP) connections"
        
        Add-UsersToSecurityGroup -GroupName "Remote Management Users" -UserList $labUsers -GroupDescription "Allows remote management via WinRM and PowerShell"
        
        # Try to add to additional groups (if they exist)
        try {
            Add-UsersToSecurityGroup -GroupName "WinRMRemoteWMIUsers__" -UserList $labUsers -GroupDescription "Allows remote access via WMI"
        } catch {
            Write-Host "⚠️ WinRMRemoteWMIUsers group not found (this is normal)" -ForegroundColor Yellow
        }
        
        # Check final group membership (selective)
        Write-Host "`n📋 Checking remote access group membership (selective):" -ForegroundColor Cyan
        
        # Check one user from each department
        $sampleUsers = @("JohnSmith", "RobertWilson", "ChrisAnderson", "LauraWhite")
        
        foreach ($username in $sampleUsers) {
            if ($labUsers -contains $username) {
                Write-Host "`n👤 User: $username" -ForegroundColor White
                
                try {
                    # Get all user groups
                    $userMembership = Get-ADUser -Identity $username -Properties MemberOf | Select-Object -ExpandProperty MemberOf
                    
                    # Check membership in key groups
                    $rdpAccess = $userMembership | Where-Object { $_ -like "*Remote Desktop Users*" }
                    $remoteManagement = $userMembership | Where-Object { $_ -like "*Remote Management Users*" }
                    $winrmAccess = $userMembership | Where-Object { $_ -like "*WinRM*" }
                    
                    if ($rdpAccess) {
                        Write-Host "  ✅ Remote Desktop Users - RDP access allowed" -ForegroundColor Green
                    } else {
                        Write-Host "  ❌ NO Remote Desktop access" -ForegroundColor Red
                    }
                    
                    if ($remoteManagement) {
                        Write-Host "  ✅ Remote Management Users - PowerShell Remoting allowed" -ForegroundColor Green
                    }
                    
                    if ($winrmAccess) {
                        Write-Host "  ✅ WinRM access configured" -ForegroundColor Green
                    }
                    
                } catch {
                    Write-Host "  ❌ Error checking user $username : $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        
        Write-Host "`n💡 Note: Only selective users checked. All $($labUsers.Count) users are configured similarly." -ForegroundColor Yellow
        Write-Host "✅ STAGE 4 COMPLETED: Remote access configured" -ForegroundColor Green
    }
} catch {
    Write-Host "❌ Critical error configuring remote access: $($_.Exception.Message)" -ForegroundColor Red
}

# ADDITIONAL STAGE: CREATE REMOTE-RDP-ACCESS GROUP
Write-Host "`n🔐 ADDITIONAL STAGE: Creating Remote-RDP-Access group..." -ForegroundColor Cyan
Write-Host "Creating special group for RDP access and adding all users" -ForegroundColor Gray

# Define parameters for group creation
$rdpGroupName = "remote-rdp-access"
$rdpGroupPath = "OU=IT,OU=Lab-Groups,$domainDN"
$rdpGroupDescription = "Group for providing RDP access to all lab users"

Write-Host "📋 Group parameters:" -ForegroundColor Yellow
Write-Host "   Group name: $rdpGroupName" -ForegroundColor White
Write-Host "   Location: $rdpGroupPath" -ForegroundColor White
Write-Host "   Description: $rdpGroupDescription" -ForegroundColor White

# Check if group exists
try {
    $existingGroup = Get-ADGroup -Identity $rdpGroupName -ErrorAction Stop
    Write-Host "⚠️ Group $rdpGroupName already exists: $($existingGroup.DistinguishedName)" -ForegroundColor Yellow
    $rdpGroup = $existingGroup
} catch {
    # Create group
    try {
        Write-Host "🔧 Creating group $rdpGroupName..." -ForegroundColor Cyan
        
        $groupParameters = @{
            Name = $rdpGroupName
            SamAccountName = $rdpGroupName
            GroupCategory = "Security"
            GroupScope = "Global"
            DisplayName = "Remote RDP Access Group"
            Description = $rdpGroupDescription
            Path = $rdpGroupPath
        }
        
        New-ADGroup @groupParameters
        Write-Host "✅ Group $rdpGroupName successfully created in IT OU" -ForegroundColor Green
        
        # Get created group for further work
        $rdpGroup = Get-ADGroup -Identity $rdpGroupName
        
    } catch {
        Write-Host "❌ ERROR creating group $rdpGroupName : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Check access rights and existence of OU=IT,OU=Lab-Groups" -ForegroundColor Red
        return
    }
}

# Get all created users from lab environment
Write-Host "`n👥 Getting list of all lab users..." -ForegroundColor Cyan

try {
    $allLabUsers = Get-ADUser -Filter * -SearchBase "OU=Lab-Groups,$domainDN" -Properties Department, Title
    
    if ($allLabUsers.Count -eq 0) {
        Write-Host "❌ Lab users not found in OU=Lab-Groups" -ForegroundColor Red
        Write-Host "Make sure user creation script was executed successfully" -ForegroundColor Red
        return
    }
    
    Write-Host "✅ Found users: $($allLabUsers.Count)" -ForegroundColor Green
    
    # Show statistics by departments
    Write-Host "`n📊 User distribution by departments:" -ForegroundColor Yellow
    $departmentStats = $allLabUsers | Group-Object Department | Sort-Object Name
    foreach ($dept in $departmentStats) {
        $deptName = if ($dept.Name) { $dept.Name } else { "Not specified" }
        Write-Host "   $deptName : $($dept.Count) users" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "❌ Error getting user list: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Add all users to remote-rdp-access group
Write-Host "`n🔧 Adding all users to group $rdpGroupName..." -ForegroundColor Cyan

$addedUsers = 0
$existingMembers = 0
$errorCount = 0

# Get current group members for checking
try {
    $currentMembers = Get-ADGroupMember -Identity $rdpGroupName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SamAccountName
    if (-not $currentMembers) { $currentMembers = @() }
} catch {
    $currentMembers = @()
}

Write-Host "📋 Current members in group: $($currentMembers.Count)" -ForegroundColor Gray

foreach ($user in $allLabUsers) {
    $username = $user.SamAccountName
    $userDisplayName = "$($user.GivenName) $($user.Surname)"
    $userDepartment = if ($user.Department) { $user.Department } else { "Not specified" }
    
    Write-Host "👤 Processing: $userDisplayName ($username) - $userDepartment" -ForegroundColor White
    
    try {
        # Check if user is already a group member
        if ($currentMembers -contains $username) {
            Write-Host "  ⚠️ Already a group member" -ForegroundColor Yellow
            $existingMembers++
        } else {
            # Add user to group
            Add-ADGroupMember -Identity $rdpGroupName -Members $username -ErrorAction Stop
            Write-Host "  ✅ Successfully added to group" -ForegroundColor Green
            $addedUsers++
        }
    } catch {
        Write-Host "  ❌ Error adding: $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }
}

# Final statistics
Write-Host "`n📊 FINAL GROUP ADDITION STATISTICS:" -ForegroundColor Cyan
Write-Host "✅ Successfully added: $addedUsers users" -ForegroundColor Green
Write-Host "⚠️ Already in group: $existingMembers users" -ForegroundColor Yellow
Write-Host "❌ Addition errors: $errorCount" -ForegroundColor Red
Write-Host "📊 Total processed: $($allLabUsers.Count) users" -ForegroundColor White

# Check final group state
Write-Host "`n🔍 Checking final state of group $rdpGroupName :" -ForegroundColor Cyan

try {
    $finalMembers = Get-ADGroupMember -Identity $rdpGroupName | Select-Object Name, SamAccountName, ObjectClass
    
    Write-Host "✅ Total members in group: $($finalMembers.Count)" -ForegroundColor Green
    
    if ($finalMembers.Count -gt 0) {
        Write-Host "`n📋 Members of group $rdpGroupName :" -ForegroundColor Yellow
        
        # Group by departments for nice display
        $membersByDept = @{}
        
        foreach ($member in $finalMembers) {
            try {
                $memberDetails = Get-ADUser -Identity $member.SamAccountName -Properties Department, Title -ErrorAction SilentlyContinue
                $dept = if ($memberDetails.Department) { $memberDetails.Department } else { "Not specified" }
                
                if (-not $membersByDept.ContainsKey($dept)) {
                    $membersByDept[$dept] = @()
                }
                
                $membersByDept[$dept] += @{
                    Name = $member.Name
                    SamAccountName = $member.SamAccountName
                    Title = if ($memberDetails.Title) { $memberDetails.Title } else { "Not specified" }
                }
            } catch {
                Write-Host "  ⚠️ Could not get details for $($member.SamAccountName)" -ForegroundColor Yellow
            }
        }
        
        # Display members by departments
        $sortedDepts = $membersByDept.Keys | Sort-Object
        foreach ($dept in $sortedDepts) {
            Write-Host "`n🏢 Department: $dept ($($membersByDept[$dept].Count) people)" -ForegroundColor Cyan
            foreach ($member in ($membersByDept[$dept] | Sort-Object Name)) {
                Write-Host "   👤 $($member.Name) ($($member.SamAccountName)) - $($member.Title)" -ForegroundColor Gray
            }
        }
    }
    
} catch {
    Write-Host "❌ Error checking group members: $($_.Exception.Message)" -ForegroundColor Red
}

# Show information about created group
Write-Host "`n📋 CREATED GROUP INFORMATION:" -ForegroundColor Cyan
try {
    $groupInfo = Get-ADGroup -Identity $rdpGroupName -Properties Description, ManagedBy, GroupCategory, GroupScope, whenCreated
    
    Write-Host "Group name: $($groupInfo.Name)" -ForegroundColor White
    Write-Host "SAM Account Name: $($groupInfo.SamAccountName)" -ForegroundColor White
    Write-Host "Distinguished Name: $($groupInfo.DistinguishedName)" -ForegroundColor White
    Write-Host "Category: $($groupInfo.GroupCategory)" -ForegroundColor White
    Write-Host "Scope: $($groupInfo.GroupScope)" -ForegroundColor White
    Write-Host "Description: $($groupInfo.Description)" -ForegroundColor White
    Write-Host "Created date: $($groupInfo.whenCreated)" -ForegroundColor White
    Write-Host "Number of members: $($finalMembers.Count)" -ForegroundColor White
    
} catch {
    Write-Host "❌ Error getting group information: $($_.Exception.Message)" -ForegroundColor Red
}

# Instructions for using the new group
Write-Host "`n💡 INSTRUCTIONS FOR USING REMOTE-RDP-ACCESS GROUP:" -ForegroundColor Yellow
Write-Host "1. Group created in OU=IT,OU=Lab-Groups for managing RDP access" -ForegroundColor White
Write-Host "2. All lab users added to this group" -ForegroundColor White
Write-Host "3. Use this group for:" -ForegroundColor White
Write-Host "   • Assigning RDP connection permissions" -ForegroundColor Gray
Write-Host "   • Managing Terminal Services access" -ForegroundColor Gray
Write-Host "   • Configuring Group Policy for remote access" -ForegroundColor Gray
Write-Host "   • Centralized RDP rights management" -ForegroundColor Gray
Write-Host "4. To add/remove RDP access - manage group membership" -ForegroundColor White
Write-Host "5. Group is a Global Security Group - can be used in any domain in forest" -ForegroundColor White

Write-Host "`n🚀 USAGE EXAMPLES:" -ForegroundColor Yellow
Write-Host "• Group Policy: Computer Configuration → Windows Settings → Security Settings → User Rights Assignment → 'Allow log on through Remote Desktop Services'" -ForegroundColor Gray
Write-Host "• Assign right to group: $($domain.DNSRoot)\$rdpGroupName" -ForegroundColor Gray
Write-Host "• PowerShell: Add-ADGroupMember -Identity '$rdpGroupName' -Members 'NewUser'" -ForegroundColor Gray
Write-Host "• PowerShell: Remove-ADGroupMember -Identity '$rdpGroupName' -Members 'UserToRemove'" -ForegroundColor Gray

Write-Host "`n✅ ADDITIONAL STAGE COMPLETED: Remote-RDP-Access group ready to use!" -ForegroundColor Green
Write-Host "   Location: OU=IT,OU=Lab-Groups,$($domain.DNSRoot)" -ForegroundColor White
Write-Host "   Members in group: $($finalMembers.Count)" -ForegroundColor White
Write-Host "   Type: Global Security Group" -ForegroundColor White

# FINAL REPORT AND INSTRUCTIONS
Write-Host "`n" + "=" * 80 -ForegroundColor Magenta
Write-Host "🎉 LAB ENVIRONMENT FOR STUDYING GROUPS IS FULLY READY!" -ForegroundColor Magenta
Write-Host "=" * 80 -ForegroundColor Magenta

Write-Host "`n📋 SUMMARY OF CREATED RESOURCES:" -ForegroundColor Cyan

# Domain summary
Write-Host "`n🌐 DOMAIN INFORMATION:" -ForegroundColor Yellow
Write-Host "   Domain name: $($domain.DNSRoot)" -ForegroundColor White
Write-Host "   Distinguished Name: $domainDN" -ForegroundColor White
Write-Host "   Domain functional level: $($domain.DomainMode)" -ForegroundColor White
Write-Host "   Forest functional level: $((Get-ADForest).ForestMode)" -ForegroundColor White
Write-Host "   Domain controller: $($env:COMPUTERNAME)" -ForegroundColor White

# User summary
try {
   $finalUserCount = (Get-ADUser -Filter * -SearchBase "OU=Lab-Groups,$domainDN").Count
   Write-Host "`n👥 USERS:" -ForegroundColor Yellow
   Write-Host "   Total created: $finalUserCount" -ForegroundColor White
   Write-Host "   Password for all: Somepass1" -ForegroundColor White
   Write-Host "   Organization: Learn IT Lessons" -ForegroundColor White
   
   # Detailed statistics by departments
   Write-Host "`n   Distribution by departments:" -ForegroundColor White
   $departmentNames | ForEach-Object {
       $deptName = $_
       try {
           $deptCount = (Get-ADUser -Filter "Department -eq '$deptName'" -SearchBase "OU=Lab-Groups,$domainDN" -Properties Department).Count
           Write-Host "     $deptName department: $deptCount users" -ForegroundColor Gray
       } catch {
           Write-Host "     $deptName department: count error" -ForegroundColor Red
       }
   }
} catch {
   Write-Host "`n👥 USERS: Count error" -ForegroundColor Red
}

# File summary
try {
   $folderCount = (Get-ChildItem "C:\LabShares" -Directory -ErrorAction SilentlyContinue).Count
   $fileCount = (Get-ChildItem "C:\LabShares" -File -Recurse -ErrorAction SilentlyContinue).Count
   Write-Host "`n📁 FILE RESOURCES:" -ForegroundColor Yellow
   Write-Host "   Main folder: C:\LabShares" -ForegroundColor White
   Write-Host "   Created folders: $folderCount" -ForegroundColor White
   Write-Host "   Created files: $fileCount" -ForegroundColor White
   Write-Host "   Department folders:" -ForegroundColor White
   Write-Host "     - IT-Resources (for IT department)" -ForegroundColor Gray
   Write-Host "     - Finance-Data (for Finance department)" -ForegroundColor Gray
   Write-Host "     - Company-Announcements (for everyone)" -ForegroundColor Gray
   Write-Host "     - Marketing-Materials (for Marketing department)" -ForegroundColor Gray
   Write-Host "     - TestFile.txt (for group conversion demonstration)" -ForegroundColor Gray
} catch {
   Write-Host "`n📁 FILE RESOURCES: Count error" -ForegroundColor Red
}

# Usage instructions
Write-Host "`n🚀 READY FOR DEMONSTRATION!" -ForegroundColor Yellow
Write-Host "`nNow you can start practical demonstrations:" -ForegroundColor White
Write-Host "✅ Creating security groups and distribution groups" -ForegroundColor Gray
Write-Host "✅ Assigning permissions to files and folders" -ForegroundColor Gray
Write-Host "✅ Demonstrating access token updates" -ForegroundColor Gray
Write-Host "✅ Converting between group types" -ForegroundColor Gray
Write-Host "✅ Testing resource access" -ForegroundColor Gray
Write-Host "✅ Working with distribution groups for email" -ForegroundColor Gray

Write-Host "`n🔐 CONNECTION INFORMATION:" -ForegroundColor Yellow
Write-Host "   Login format: $($domain.DNSRoot)\Username" -ForegroundColor White
Write-Host "   Example logins by departments:" -ForegroundColor White
Write-Host "     IT department:" -ForegroundColor Cyan
Write-Host "       $($domain.DNSRoot)\JohnSmith (System Administrator)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\MaryJohnson (Developer)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\MichaelBrown (Network Administrator)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\SarahDavis (QA Tester)" -ForegroundColor Gray
Write-Host "     Finance department:" -ForegroundColor Cyan
Write-Host "       $($domain.DNSRoot)\RobertWilson (Chief Accountant)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\LisaMiller (Financial Analyst)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\DavidMoore (Economist)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\JenniferTaylor (Auditor)" -ForegroundColor Gray
Write-Host "     Marketing department:" -ForegroundColor Cyan
Write-Host "       $($domain.DNSRoot)\ChrisAnderson (Marketing Specialist)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\AmandaThomas (PR Manager)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\KevinJackson (SMM Specialist)" -ForegroundColor Gray
Write-Host "     HR department:" -ForegroundColor Cyan
Write-Host "       $($domain.DNSRoot)\LauraWhite (HR Specialist)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\EmilyMartin (Recruiter)" -ForegroundColor Gray
Write-Host "       $($domain.DNSRoot)\JamesGarcia (Personnel Manager)" -ForegroundColor Gray
Write-Host "   Password for all: Somepass1" -ForegroundColor White

Write-Host "`n🛠️ AVAILABLE CONNECTION METHODS:" -ForegroundColor Yellow
Write-Host "   ✅ Remote Desktop Protocol (RDP)" -ForegroundColor Green
Write-Host "   ✅ PowerShell Remoting" -ForegroundColor Green
Write-Host "   ✅ Windows Remote Management (WinRM)" -ForegroundColor Green

Write-Host "`n📚 STEP-BY-STEP DEMONSTRATION GUIDE:" -ForegroundColor Yellow
Write-Host "1. Open Active Directory Users and Computers (dsa.msc)" -ForegroundColor White
Write-Host "2. Navigate to OU=Lab-Groups,DC=$($domain.DNSRoot.Replace('.', ',DC='))" -ForegroundColor White
Write-Host "3. Create security groups:" -ForegroundColor White
Write-Host "   - IT-Security (for IT department)" -ForegroundColor Gray
Write-Host "   - Finance-Security (for Finance department)" -ForegroundColor Gray
Write-Host "4. Create distribution groups:" -ForegroundColor White
Write-Host "   - All-Employees (for all employees)" -ForegroundColor Gray
Write-Host "   - IT-Announcements (for IT notifications)" -ForegroundColor Gray
Write-Host "5. Add users to appropriate groups" -ForegroundColor White
Write-Host "6. Configure permissions on folders in C:\LabShares" -ForegroundColor White
Write-Host "7. Test user access" -ForegroundColor White
Write-Host "8. Demonstrate group conversion and its impact" -ForegroundColor White

Write-Host "`n💡 DEMONSTRATION RECOMMENDATIONS:" -ForegroundColor Yellow
Write-Host "• Use users from different departments for testing" -ForegroundColor White
Write-Host "• Demonstrate differences in security group and distribution group behavior" -ForegroundColor White
Write-Host "• Show importance of re-login after group membership changes" -ForegroundColor White
Write-Host "• Use 'whoami /groups' command to check access tokens" -ForegroundColor White
Write-Host "• Demonstrate Exchange Server functionality (if available)" -ForegroundColor White

Write-Host "`n🎯 DEMONSTRATION SCENARIOS:" -ForegroundColor Yellow
Write-Host "Scenario 1: IT department gets access to IT-Resources" -ForegroundColor White
Write-Host "  - Create IT-Security group" -ForegroundColor Gray
Write-Host "  - Add JohnSmith, MaryJohnson, MichaelBrown, SarahDavis" -ForegroundColor Gray
Write-Host "  - Assign permissions to C:\LabShares\IT-Resources" -ForegroundColor Gray
Write-Host "  - Test access" -ForegroundColor Gray

Write-Host "`nScenario 2: Finance department and confidential data" -ForegroundColor White
Write-Host "  - Create Finance-Security group" -ForegroundColor Gray
Write-Host "  - Add RobertWilson, LisaMiller, DavidMoore, JenniferTaylor" -ForegroundColor Gray
Write-Host "  - Assign permissions to C:\LabShares\Finance-Data" -ForegroundColor Gray
Write-Host "  - Verify other departments don't have access" -ForegroundColor Gray

Write-Host "`nScenario 3: Distribution groups for announcements" -ForegroundColor White
Write-Host "  - Create All-Employees group (type: Distribution)" -ForegroundColor Gray
Write-Host "  - Add all users" -ForegroundColor Gray
Write-Host "  - Try to assign file permissions (should not work)" -ForegroundColor Gray
Write-Host "  - Show usage for email distribution" -ForegroundColor Gray

Write-Host "`nScenario 4: Group conversion" -ForegroundColor White
Write-Host "  - Create group, assign permissions to TestFile.txt" -ForegroundColor Gray
Write-Host "  - Convert to Distribution group" -ForegroundColor Gray
Write-Host "  - Show loss of access" -ForegroundColor Gray
Write-Host "  - Convert back to Security group" -ForegroundColor Gray
Write-Host "  - Show access restoration" -ForegroundColor Gray

Write-Host "`n" + "=" * 80 -ForegroundColor Magenta
Write-Host "🎓 ENJOY YOUR LEARNING! LAB ENVIRONMENT IS READY TO USE!" -ForegroundColor Magenta
Write-Host "   Total users: $($usersToCreate.Count)" -ForegroundColor White
Write-Host "   IT: 4 | Finance: 4 | Marketing: 3 | HR: 3" -ForegroundColor White
Write-Host "=" * 80 -ForegroundColor Magenta

# End of script - save execution log
$logPath = "C:\LabShares\Lab-Setup-Log-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
try {
   $logContent = @"
WINDOWS SERVER 2025 GROUPS LAB ENVIRONMENT SETUP LOG
=====================================================

Execution date: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')
Domain: $($domain.DNSRoot)
Distinguished Name: $domainDN
Functional level: $($domain.DomainMode)

CREATED USERS ($($usersToCreate.Count) people):
========================================

IT DEPARTMENT (4 people):
- JohnSmith (John Smith) - System Administrator
- MaryJohnson (Mary Johnson) - Developer  
- MichaelBrown (Michael Brown) - Network Administrator
- SarahDavis (Sarah Davis) - QA Tester

FINANCE DEPARTMENT (4 people):
- RobertWilson (Robert Wilson) - Chief Accountant
- LisaMiller (Lisa Miller) - Financial Analyst
- DavidMoore (David Moore) - Economist
- JenniferTaylor (Jennifer Taylor) - Auditor

MARKETING DEPARTMENT (3 people):
- ChrisAnderson (Chris Anderson) - Marketing Specialist
- AmandaThomas (Amanda Thomas) - PR Manager
- KevinJackson (Kevin Jackson) - SMM Specialist

HR DEPARTMENT (3 people):
- LauraWhite (Laura White) - HR Specialist
- EmilyMartin (Emily Martin) - Recruiter
- JamesGarcia (James Garcia) - Personnel Manager

CREATED RESOURCES:
==================
OU structure: OU=Lab-Groups with IT, Finance, Marketing, HR subdivisions
Folders: C:\LabShares\IT-Resources, Finance-Data, Company-Announcements, Marketing-Materials
Files: TestFile.txt and corresponding files in each folder

ACCESS SETTINGS:
================
All users added to Remote Desktop Users and Remote Management Users groups
Password for all users: Somepass1

LOGIN FORMAT:
=============
$($domain.DNSRoot)\Username
Examples: $($domain.DNSRoot)\JohnSmith, $($domain.DNSRoot)\RobertWilson

STATUS: Script executed successfully
READINESS: Lab environment ready for groups demonstration

© Learn IT Lessons - Windows Server 2025 Lab Environment
"@
   
   $logContent | Out-File $logPath -Encoding UTF8
   Write-Host "`n📝 Detailed log saved: $logPath" -ForegroundColor Cyan
   Write-Host "   Contains complete information about created users and resources" -ForegroundColor Gray
} catch {
   Write-Host "`n⚠️ Could not save execution log: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n🔚 SCRIPT COMPLETED SUCCESSFULLY" -ForegroundColor Green
Write-Host "Execution time: $(Get-Date)" -ForegroundColor Gray
