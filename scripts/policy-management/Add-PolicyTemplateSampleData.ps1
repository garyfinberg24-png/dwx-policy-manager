# Populate Policy Templates and Metadata with Sample Data
# This script adds realistic sample data for policy authoring

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

# Connect to SharePoint
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "Populating Policy Templates with sample data..." -ForegroundColor Cyan

# ============================================================================
# 1. Policy Templates - Company-Approved Templates
# ============================================================================
Write-Host "Adding sample policy templates..." -ForegroundColor Yellow

$templates = @(
    @{
        Title = "Standard HR Policy Template"
        TemplateType = "HR Policy"
        TemplateCategory = "HR"
        TemplateDescription = "Standard template for HR policies including sections for purpose, scope, policy statement, procedures, and responsibilities."
        TemplateContent = @"
<h1>[Policy Name]</h1>
<h2>1. Purpose</h2>
<p>This policy establishes [describe purpose]...</p>

<h2>2. Scope</h2>
<p>This policy applies to [describe scope]...</p>

<h2>3. Policy Statement</h2>
<p>[Company Name] is committed to [state policy]...</p>

<h2>4. Definitions</h2>
<ul>
<li><strong>Term 1:</strong> Definition...</li>
<li><strong>Term 2:</strong> Definition...</li>
</ul>

<h2>5. Procedures</h2>
<ol>
<li>Step 1...</li>
<li>Step 2...</li>
<li>Step 3...</li>
</ol>

<h2>6. Responsibilities</h2>
<p><strong>Employees:</strong> [Describe responsibilities]</p>
<p><strong>Managers:</strong> [Describe responsibilities]</p>
<p><strong>HR:</strong> [Describe responsibilities]</p>

<h2>7. Non-Compliance</h2>
<p>Failure to comply with this policy may result in [consequences]...</p>

<h2>8. Related Policies</h2>
<ul>
<li>Policy A</li>
<li>Policy B</li>
</ul>

<h2>9. Contact Information</h2>
<p>For questions, contact: [Department/Person]</p>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "Medium"
        SuggestedReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        KeyPointsTemplate = "Must read by [timeframe];Must acknowledge understanding;Must follow procedures;Contact HR with questions"
        Tags = "HR;Standard;Template;General"
    },
    @{
        Title = "IT Security Policy Template"
        TemplateType = "IT Policy"
        TemplateCategory = "IT"
        TemplateDescription = "Template for IT security policies covering access control, data protection, and security best practices."
        TemplateContent = @"
<h1>[IT Security Policy Name]</h1>
<h2>1. Overview</h2>
<p>This policy defines security requirements for [describe area]...</p>

<h2>2. Purpose</h2>
<p>To ensure the confidentiality, integrity, and availability of [company] information assets.</p>

<h2>3. Scope</h2>
<p>Applies to all employees, contractors, and third parties with access to [company] systems.</p>

<h2>4. Security Requirements</h2>
<h3>4.1 Access Control</h3>
<ul>
<li>Use strong passwords (minimum 12 characters)</li>
<li>Enable multi-factor authentication</li>
<li>Do not share credentials</li>
</ul>

<h3>4.2 Data Protection</h3>
<ul>
<li>Encrypt sensitive data at rest and in transit</li>
<li>Use approved cloud storage only</li>
<li>Follow data classification guidelines</li>
</ul>

<h3>4.3 Device Security</h3>
<ul>
<li>Keep software updated</li>
<li>Use company-approved antivirus</li>
<li>Report lost/stolen devices immediately</li>
</ul>

<h2>5. Incident Reporting</h2>
<p>Report security incidents to IT Security within 1 hour: security@company.com</p>

<h2>6. Consequences</h2>
<p>Violations may result in disciplinary action up to and including termination.</p>

<h2>7. Review</h2>
<p>This policy is reviewed annually and updated as needed.</p>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "High"
        SuggestedReadTimeframe = "Day 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        KeyPointsTemplate = "Use strong passwords;Enable MFA;Encrypt sensitive data;Report incidents immediately;Keep devices updated"
        Tags = "IT;Security;High Risk;Mandatory"
    },
    @{
        Title = "Code of Conduct Template"
        TemplateType = "Code of Conduct"
        TemplateCategory = "General"
        TemplateDescription = "Standard template for organizational code of conduct and ethical behavior expectations."
        TemplateContent = @"
<h1>Code of Conduct</h1>
<h2>1. Introduction</h2>
<p>[Company Name] is committed to maintaining the highest standards of ethics and integrity.</p>

<h2>2. Our Values</h2>
<ul>
<li><strong>Integrity:</strong> We act honestly and ethically</li>
<li><strong>Respect:</strong> We treat everyone with dignity</li>
<li><strong>Accountability:</strong> We take responsibility for our actions</li>
<li><strong>Excellence:</strong> We strive for quality in all we do</li>
</ul>

<h2>3. Expected Behaviors</h2>
<h3>3.1 Professional Conduct</h3>
<p>Employees must maintain professional behavior at all times...</p>

<h3>3.2 Respect and Diversity</h3>
<p>We celebrate diversity and promote an inclusive workplace...</p>

<h3>3.3 Conflicts of Interest</h3>
<p>Avoid situations where personal interests conflict with company interests...</p>

<h2>4. Prohibited Conduct</h2>
<ul>
<li>Harassment or discrimination</li>
<li>Misuse of company resources</li>
<li>Bribery or corruption</li>
<li>Disclosure of confidential information</li>
</ul>

<h2>5. Reporting Concerns</h2>
<p>Report violations via:</p>
<ul>
<li>Direct supervisor</li>
<li>HR Department</li>
<li>Ethics Hotline: 1-800-XXX-XXXX</li>
</ul>

<h2>6. Protection Against Retaliation</h2>
<p>We prohibit retaliation against anyone who reports concerns in good faith.</p>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "High"
        SuggestedReadTimeframe = "Day 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        KeyPointsTemplate = "Act with integrity;Treat others with respect;Avoid conflicts of interest;Report violations;No retaliation"
        Tags = "Ethics;Code of Conduct;Mandatory;High Risk"
    },
    @{
        Title = "Data Privacy Policy Template"
        TemplateType = "Data Privacy"
        TemplateCategory = "Data Privacy"
        TemplateDescription = "GDPR-compliant template for data privacy and protection policies."
        TemplateContent = @"
<h1>Data Privacy and Protection Policy</h1>
<h2>1. Purpose</h2>
<p>To ensure compliance with data protection laws including GDPR, CCPA, and other regulations.</p>

<h2>2. Scope</h2>
<p>Applies to all processing of personal data by [Company].</p>

<h2>3. Data Protection Principles</h2>
<ul>
<li><strong>Lawfulness:</strong> Process data legally and transparently</li>
<li><strong>Purpose Limitation:</strong> Collect for specific purposes</li>
<li><strong>Data Minimization:</strong> Collect only what is necessary</li>
<li><strong>Accuracy:</strong> Keep data accurate and up to date</li>
<li><strong>Storage Limitation:</strong> Retain only as long as necessary</li>
<li><strong>Security:</strong> Protect against unauthorized access</li>
</ul>

<h2>4. Individual Rights</h2>
<ul>
<li>Right to access</li>
<li>Right to rectification</li>
<li>Right to erasure</li>
<li>Right to data portability</li>
<li>Right to object</li>
</ul>

<h2>5. Data Processing Requirements</h2>
<h3>5.1 Collection</h3>
<p>Obtain consent before collecting personal data...</p>

<h3>5.2 Storage</h3>
<p>Store data securely using encryption...</p>

<h3>5.3 Sharing</h3>
<p>Only share data with authorized parties...</p>

<h2>6. Data Breach Response</h2>
<p>Report breaches to DPO within 24 hours: dpo@company.com</p>

<h2>7. Training</h2>
<p>All employees must complete annual data privacy training.</p>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "High"
        SuggestedReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        KeyPointsTemplate = "Follow data protection principles;Respect individual rights;Obtain consent;Secure data;Report breaches promptly"
        Tags = "GDPR;Privacy;Compliance;High Risk"
    },
    @{
        Title = "Health and Safety Policy Template"
        TemplateType = "Health & Safety"
        TemplateCategory = "Health & Safety"
        TemplateDescription = "Template for workplace health and safety policies."
        TemplateContent = @"
<h1>Health and Safety Policy</h1>
<h2>1. Policy Statement</h2>
<p>[Company] is committed to providing a safe and healthy workplace for all employees.</p>

<h2>2. Objectives</h2>
<ul>
<li>Prevent workplace injuries and illnesses</li>
<li>Comply with health and safety regulations</li>
<li>Promote safety awareness and training</li>
</ul>

<h2>3. Responsibilities</h2>
<h3>3.1 Management</h3>
<ul>
<li>Provide safe work environment</li>
<li>Conduct risk assessments</li>
<li>Provide safety training</li>
</ul>

<h3>3.2 Employees</h3>
<ul>
<li>Follow safety procedures</li>
<li>Use PPE when required</li>
<li>Report hazards immediately</li>
</ul>

<h2>4. Emergency Procedures</h2>
<h3>4.1 Fire</h3>
<p>Activate alarm, evacuate via nearest exit, assemble at [location]</p>

<h3>4.2 Medical Emergency</h3>
<p>Call emergency services, notify first aider, provide assistance</p>

<h2>5. Incident Reporting</h2>
<p>Report all incidents, near misses, and hazards to supervisor immediately.</p>

<h2>6. Training</h2>
<p>All employees receive:</p>
<ul>
<li>Induction safety training</li>
<li>Job-specific safety training</li>
<li>Annual refresher training</li>
</ul>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "High"
        SuggestedReadTimeframe = "Day 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        KeyPointsTemplate = "Follow safety procedures;Use PPE;Report hazards;Know emergency procedures;Complete safety training"
        Tags = "Health & Safety;Mandatory;Compliance"
    },
    @{
        Title = "Remote Work Policy Template"
        TemplateType = "HR Policy"
        TemplateCategory = "HR"
        TemplateDescription = "Template for remote and hybrid work arrangements."
        TemplateContent = @"
<h1>Remote Work Policy</h1>
<h2>1. Purpose</h2>
<p>To establish guidelines for remote work arrangements while maintaining productivity and security.</p>

<h2>2. Eligibility</h2>
<p>Remote work is available to employees whose roles can be performed effectively off-site.</p>

<h2>3. Work Arrangements</h2>
<h3>3.1 Fully Remote</h3>
<p>Work from home or approved location full-time</p>

<h3>3.2 Hybrid</h3>
<p>Split time between office and remote location</p>

<h2>4. Requirements</h2>
<ul>
<li>Dedicated workspace</li>
<li>Reliable internet connection</li>
<li>Availability during core hours</li>
<li>Regular communication with team</li>
</ul>

<h2>5. Equipment and Expenses</h2>
<p>Company provides:</p>
<ul>
<li>Laptop and necessary peripherals</li>
<li>Software licenses</li>
<li>Monthly internet stipend (if applicable)</li>
</ul>

<h2>6. Security</h2>
<ul>
<li>Use VPN for accessing company resources</li>
<li>Lock devices when unattended</li>
<li>Maintain confidentiality</li>
</ul>

<h2>7. Performance and Communication</h2>
<ul>
<li>Attend scheduled meetings</li>
<li>Maintain regular communication</li>
<li>Meet deadlines and deliverables</li>
</ul>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "Medium"
        SuggestedReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        KeyPointsTemplate = "Maintain dedicated workspace;Use VPN;Stay available during core hours;Communicate regularly;Protect company data"
        Tags = "HR;Remote Work;Flexible Work"
    },
    @{
        Title = "Expense Reimbursement Policy Template"
        TemplateType = "Financial Policy"
        TemplateCategory = "Finance"
        TemplateDescription = "Template for employee expense reporting and reimbursement."
        TemplateContent = @"
<h1>Expense Reimbursement Policy</h1>
<h2>1. Purpose</h2>
<p>To establish guidelines for reimbursing legitimate business expenses.</p>

<h2>2. Eligible Expenses</h2>
<ul>
<li>Travel (flights, hotels, ground transportation)</li>
<li>Meals (within per diem limits)</li>
<li>Client entertainment (with approval)</li>
<li>Office supplies</li>
<li>Professional development</li>
</ul>

<h2>3. Expense Limits</h2>
<h3>3.1 Meals</h3>
<ul>
<li>Breakfast: $15</li>
<li>Lunch: $25</li>
<li>Dinner: $50</li>
</ul>

<h3>3.2 Hotels</h3>
<p>Up to $200/night (varies by location)</p>

<h2>4. Submission Requirements</h2>
<ul>
<li>Submit within 30 days</li>
<li>Provide itemized receipts</li>
<li>Include business purpose</li>
<li>Obtain manager approval</li>
</ul>

<h2>5. Non-Reimbursable Expenses</h2>
<ul>
<li>Personal expenses</li>
<li>Alcohol (unless client entertainment)</li>
<li>Traffic violations</li>
<li>Entertainment for personal guests</li>
</ul>

<h2>6. Reimbursement Timeline</h2>
<p>Approved expenses reimbursed within 2 pay periods.</p>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "Medium"
        SuggestedReadTimeframe = "Week 2"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        KeyPointsTemplate = "Know expense limits;Keep receipts;Submit within 30 days;Get manager approval;Understand non-reimbursable items"
        Tags = "Finance;Expenses;Reimbursement"
    },
    @{
        Title = "Social Media Policy Template"
        TemplateType = "IT Policy"
        TemplateCategory = "IT"
        TemplateDescription = "Template for social media usage and company representation guidelines."
        TemplateContent = @"
<h1>Social Media Policy</h1>
<h2>1. Purpose</h2>
<p>To provide guidelines for appropriate social media use by employees.</p>

<h2>2. Scope</h2>
<p>Applies to all social media platforms including LinkedIn, Twitter, Facebook, Instagram, etc.</p>

<h2>3. Personal Use Guidelines</h2>
<ul>
<li>Personal opinions are your own</li>
<li>Do not speak on behalf of company without authorization</li>
<li>Do not share confidential information</li>
<li>Maintain professional standards</li>
</ul>

<h2>4. Company Representation</h2>
<p>Only authorized personnel may post as official company representatives.</p>

<h2>5. Best Practices</h2>
<ul>
<li>Be respectful and professional</li>
<li>Protect company reputation</li>
<li>Respect copyright and intellectual property</li>
<li>Think before you post</li>
</ul>

<h2>6. Prohibited Activities</h2>
<ul>
<li>Sharing confidential information</li>
<li>Harassing or discriminatory posts</li>
<li>Making false claims</li>
<li>Engaging in illegal activities</li>
</ul>

<h2>7. Consequences</h2>
<p>Violations may result in disciplinary action.</p>
"@
        IsActive = $true
        UsageCount = 0
        ComplianceRisk = "Medium"
        SuggestedReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        KeyPointsTemplate = "Personal opinions are your own;Don't share confidential info;Be professional;Protect company reputation;Think before posting"
        Tags = "Social Media;IT;Communication"
    }
)

$templateCount = 0
foreach ($template in $templates) {
    try {
        Add-PnPListItem -List "JML_PolicyTemplates" -Values $template | Out-Null
        $templateCount++
        Write-Host "  ✓ Added: $($template.Title)" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ Failed to add: $($template.Title)" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "  Added $templateCount policy templates" -ForegroundColor Green

# ============================================================================
# 2. Policy Metadata Profiles
# ============================================================================
Write-Host "Adding sample metadata profiles..." -ForegroundColor Yellow

$metadataProfiles = @(
    @{
        Title = "Standard HR Policy Profile"
        ProfileName = "Standard HR Policy"
        PolicyCategory = "HR"
        ComplianceRisk = "Medium"
        ReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        TargetDepartments = "All Departments"
        TargetRoles = "All Employees"
        IsActive = $true
    },
    @{
        Title = "High-Risk IT Security Profile"
        ProfileName = "IT Security - High Risk"
        PolicyCategory = "IT"
        ComplianceRisk = "High"
        ReadTimeframe = "Day 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        TargetDepartments = "IT,Engineering,Operations"
        TargetRoles = "All Employees"
        IsActive = $true
    },
    @{
        Title = "Financial Policy Profile"
        ProfileName = "Financial Compliance"
        PolicyCategory = "Finance"
        ComplianceRisk = "High"
        ReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        TargetDepartments = "Finance,Accounting,Procurement"
        TargetRoles = "Finance Team,Managers,Executives"
        IsActive = $true
    },
    @{
        Title = "Data Privacy Profile"
        ProfileName = "GDPR Compliance"
        PolicyCategory = "Data Privacy"
        ComplianceRisk = "High"
        ReadTimeframe = "Day 3"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        TargetDepartments = "All Departments"
        TargetRoles = "All Employees"
        IsActive = $true
    },
    @{
        Title = "General Office Policy Profile"
        ProfileName = "General Office"
        PolicyCategory = "General"
        ComplianceRisk = "Low"
        ReadTimeframe = "Week 2"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        TargetDepartments = "All Departments"
        TargetRoles = "All Employees"
        IsActive = $true
    }
)

$profileCount = 0
foreach ($profile in $metadataProfiles) {
    try {
        Add-PnPListItem -List "JML_PolicyMetadataProfiles" -Values $profile | Out-Null
        $profileCount++
        Write-Host "  ✓ Added: $($profile.ProfileName)" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ Failed to add: $($profile.ProfileName)" -ForegroundColor Red
    }
}

Write-Host "  Added $profileCount metadata profiles" -ForegroundColor Green

# ============================================================================
# Summary
# ============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Sample Data Population Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "✓ $templateCount Policy Templates" -ForegroundColor Green
Write-Host "✓ $profileCount Metadata Profiles" -ForegroundColor Green
Write-Host ""
Write-Host "Templates include:" -ForegroundColor Yellow
Write-Host "  - HR Policies" -ForegroundColor White
Write-Host "  - IT Security" -ForegroundColor White
Write-Host "  - Code of Conduct" -ForegroundColor White
Write-Host "  - Data Privacy (GDPR)" -ForegroundColor White
Write-Host "  - Health & Safety" -ForegroundColor White
Write-Host "  - Remote Work" -ForegroundColor White
Write-Host "  - Expense Reimbursement" -ForegroundColor White
Write-Host "  - Social Media" -ForegroundColor White
Write-Host ""
Write-Host "Ready to use in Policy Author web part!" -ForegroundColor Green
Write-Host ""
