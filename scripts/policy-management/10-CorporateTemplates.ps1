# ============================================================================
# PM_CorporateTemplates - Corporate Template Library Provisioning
# Creates the document library for branded corporate policy templates
# and uploads starter template files (Word, Excel, PowerPoint).
# ============================================================================
# Prerequisites: Already connected to SharePoint via Connect-PnPOnline
# Site: https://mf7m.sharepoint.com/sites/PolicyManager
# ============================================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
$LibraryName = "PM_CorporateTemplates"
$LibraryTitle = "Corporate Templates"

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " Creating $LibraryName Document Library" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

# --------------------------------------------------
# 1. Create the document library
# --------------------------------------------------
$existingList = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "[SKIP] $LibraryName already exists." -ForegroundColor Yellow
} else {
    Write-Host "[CREATE] Creating document library: $LibraryName..." -ForegroundColor Green
    New-PnPList -Title $LibraryTitle -Url $LibraryName -Template DocumentLibrary -EnableVersioning
    Write-Host "[OK] Document library created." -ForegroundColor Green
}

# --------------------------------------------------
# 2. Add custom metadata columns
# --------------------------------------------------
$fields = @(
    @{ Name = "TemplateType"; Type = "Choice"; Choices = @("Word", "Excel", "PowerPoint", "Image", "Other"); Default = "Word" },
    @{ Name = "Description"; Type = "Note" },
    @{ Name = "Category"; Type = "Choice"; Choices = @("Corporate", "General", "Department", "Compliance", "HR", "IT", "Finance"); Default = "Corporate" },
    @{ Name = "IsDefault"; Type = "Boolean"; Default = "0" },
    @{ Name = "IsActive"; Type = "Boolean"; Default = "1" },
    @{ Name = "SortOrder"; Type = "Number" }
)

foreach ($field in $fields) {
    $existing = Get-PnPField -List $LibraryName -Identity $field.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "[SKIP] Field '$($field.Name)' already exists." -ForegroundColor Yellow
        continue
    }

    switch ($field.Type) {
        "Choice" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Choice -Choices $field.Choices -AddToDefaultView
            Write-Host "[OK] Added Choice field: $($field.Name)" -ForegroundColor Green
        }
        "Text" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Text -AddToDefaultView
            Write-Host "[OK] Added Text field: $($field.Name)" -ForegroundColor Green
        }
        "Note" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Note -AddToDefaultView
            Write-Host "[OK] Added Note field: $($field.Name)" -ForegroundColor Green
        }
        "Number" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Number -AddToDefaultView
            Write-Host "[OK] Added Number field: $($field.Name)" -ForegroundColor Green
        }
        "Boolean" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Boolean -AddToDefaultView
            Write-Host "[OK] Added Boolean field: $($field.Name)" -ForegroundColor Green
        }
    }
}

# --------------------------------------------------
# 3. Upload starter template files
# --------------------------------------------------
Write-Host "`n--- Uploading Starter Templates ---" -ForegroundColor Cyan

# Create temp directory for template files
$tempDir = Join-Path $env:TEMP "PM_CorporateTemplates"
if (!(Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir | Out-Null }

# Helper: Create a minimal valid DOCX
function New-MinimalDocx {
    param([string]$FilePath, [string]$Title, [string]$BodyText)

    # A DOCX is a ZIP containing XML files
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path $FilePath) { Remove-Item $FilePath -Force }

    $zip = [System.IO.Compression.ZipFile]::Open($FilePath, 'Create')

    # [Content_Types].xml
    $entry = $zip.CreateEntry('[Content_Types].xml')
    $writer = New-Object System.IO.StreamWriter($entry.Open())
    $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    $writer.Close()

    # _rels/.rels
    $entry = $zip.CreateEntry('_rels/.rels')
    $writer = New-Object System.IO.StreamWriter($entry.Open())
    $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    $writer.Close()

    # word/_rels/document.xml.rels
    $entry = $zip.CreateEntry('word/_rels/document.xml.rels')
    $writer = New-Object System.IO.StreamWriter($entry.Open())
    $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    $writer.Close()

    # word/document.xml — the actual content
    $entry = $zip.CreateEntry('word/document.xml')
    $writer = New-Object System.IO.StreamWriter($entry.Open())
    $ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    $writer.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>")
    $writer.Write("<w:document xmlns:w=`"$ns`"><w:body>")
    # Title paragraph (bold, large)
    $writer.Write("<w:p><w:pPr><w:jc w:val=`"center`"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val=`"48`"/></w:rPr><w:t>$Title</w:t></w:r></w:p>")
    # Separator
    $writer.Write("<w:p><w:pPr><w:jc w:val=`"center`"/></w:pPr><w:r><w:rPr><w:color w:val=`"0D9488`"/><w:sz w:val=`"20`"/></w:rPr><w:t>DWx Policy Manager - Corporate Template</w:t></w:r></w:p>")
    $writer.Write("<w:p/>")
    # Body text
    foreach ($line in ($BodyText -split "`n")) {
        $writer.Write("<w:p><w:r><w:t xml:space=`"preserve`">$line</w:t></w:r></w:p>")
    }
    $writer.Write("</w:body></w:document>")
    $writer.Close()

    $zip.Dispose()
}

# Helper: Create a minimal valid XLSX
function New-MinimalXlsx {
    param([string]$FilePath, [string]$Title)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path $FilePath) { Remove-Item $FilePath -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($FilePath, 'Create')

    $entry = $zip.CreateEntry('[Content_Types].xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
    $w.Close()

    $entry = $zip.CreateEntry('_rels/.rels')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    $w.Close()

    $entry = $zip.CreateEntry('xl/_rels/workbook.xml.rels')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
    $w.Close()

    $entry = $zip.CreateEntry('xl/workbook.xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Policy Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
    $w.Close()

    $entry = $zip.CreateEntry('xl/sharedStrings.xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6"><si><t>Policy Name</t></si><si><t>Category</t></si><si><t>Status</t></si><si><t>Risk Level</t></si><si><t>Effective Date</t></si><si><t>Owner</t></si></sst>')
    $w.Close()

    $entry = $zip.CreateEntry('xl/worksheets/sheet1.xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1" t="s"><v>4</v></c><c r="F1" t="s"><v>5</v></c></row></sheetData></worksheet>')
    $w.Close()

    $zip.Dispose()
}

# Helper: Create a minimal valid PPTX
function New-MinimalPptx {
    param([string]$FilePath, [string]$Title)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path $FilePath) { Remove-Item $FilePath -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($FilePath, 'Create')

    $entry = $zip.CreateEntry('[Content_Types].xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>')
    $w.Close()

    $entry = $zip.CreateEntry('_rels/.rels')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/_rels/presentation.xml.rels')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/presentation.xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/slideMasters/_rels/slideMaster1.xml.rels')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/slideMasters/slideMaster1.xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/slideLayouts/_rels/slideLayout1.xml.rels')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/slideLayouts/slideLayout1.xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/slides/_rels/slide1.xml.rels')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
    $w.Close()

    $entry = $zip.CreateEntry('ppt/slides/slide1.xml')
    $w = New-Object System.IO.StreamWriter($entry.Open())
    $w.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sld xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:p=`"http://schemas.openxmlformats.org/presentationml/2006/main`"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=`"1`" name=`"`"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id=`"2`" name=`"Title`"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=`"1524000`" y=`"1397000`"/><a:ext cx=`"9144000`" cy=`"2387600`"/></a:xfrm><a:prstGeom prst=`"rect`"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor=`"ctr`"/><a:lstStyle/><a:p><a:pPr algn=`"ctr`"/><a:r><a:rPr lang=`"en-US`" sz=`"4400`" b=`"1`"/><a:t>$Title</a:t></a:r></a:p><a:p><a:pPr algn=`"ctr`"/><a:r><a:rPr lang=`"en-US`" sz=`"2000`"><a:solidFill><a:srgbClr val=`"0D9488`"/></a:solidFill></a:rPr><a:t>DWx Policy Manager - Corporate Template</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>")
    $w.Close()

    $zip.Dispose()
}

# --- Template 1: Corporate Policy Standard (Word) ---
$docxPath = Join-Path $tempDir "Corporate-Policy-Standard.docx"
New-MinimalDocx -FilePath $docxPath -Title "Corporate Policy Template" -BodyText @"
1. PURPOSE
[Describe the purpose and scope of this policy]

2. SCOPE
[Define who this policy applies to and under what circumstances]

3. POLICY STATEMENT
[State the official policy position]

4. RESPONSIBILITIES
[List roles and their responsibilities under this policy]

5. COMPLIANCE
[Describe compliance requirements and consequences of non-compliance]

6. REVIEW
[State the review cycle and responsible parties]

7. APPROVAL
Prepared by: ___________________    Date: ___________
Approved by: ___________________    Date: ___________
"@

$existingFile = Get-PnPFile -Url "$LibraryName/Corporate-Policy-Standard.docx" -ErrorAction SilentlyContinue
if ($existingFile) {
    Write-Host "[SKIP] Corporate-Policy-Standard.docx already exists." -ForegroundColor Yellow
} else {
    Add-PnPFile -Path $docxPath -Folder $LibraryName -Values @{
        Title = "Corporate Policy - Standard A4"
        TemplateType = "Word"
        Description = "Standard corporate policy template with sections for Purpose, Scope, Policy Statement, Responsibilities, Compliance, Review, and Approval."
        Category = "Corporate"
        IsDefault = $true
        IsActive = $true
        SortOrder = 1
    }
    Write-Host "[OK] Uploaded: Corporate-Policy-Standard.docx" -ForegroundColor Green
}

# --- Template 2: Executive Brief (Word) ---
$docxPath2 = Join-Path $tempDir "Corporate-Executive-Brief.docx"
New-MinimalDocx -FilePath $docxPath2 -Title "Executive Policy Brief" -BodyText @"
EXECUTIVE SUMMARY
[Concise summary of the policy for board and executive review]

KEY DECISIONS REQUIRED
- Decision 1: [Description]
- Decision 2: [Description]

IMPACT ASSESSMENT
[Business impact, resource requirements, timeline]

RISK ANALYSIS
[Key risks and mitigation strategies]

RECOMMENDATION
[Recommended course of action]

ACTION ITEMS
[ ] Action 1 - Owner - Due Date
[ ] Action 2 - Owner - Due Date
"@

$existingFile = Get-PnPFile -Url "$LibraryName/Corporate-Executive-Brief.docx" -ErrorAction SilentlyContinue
if ($existingFile) {
    Write-Host "[SKIP] Corporate-Executive-Brief.docx already exists." -ForegroundColor Yellow
} else {
    Add-PnPFile -Path $docxPath2 -Folder $LibraryName -Values @{
        Title = "Corporate Policy - Executive Brief"
        TemplateType = "Word"
        Description = "Condensed executive briefing format for board-level policies. Includes executive summary, key decisions, impact assessment, and action items."
        Category = "Corporate"
        IsDefault = $false
        IsActive = $true
        SortOrder = 2
    }
    Write-Host "[OK] Uploaded: Corporate-Executive-Brief.docx" -ForegroundColor Green
}

# --- Template 3: General Department Policy (Word) ---
$docxPath3 = Join-Path $tempDir "General-Department-Policy.docx"
New-MinimalDocx -FilePath $docxPath3 -Title "Department Policy" -BodyText @"
DEPARTMENT: [Department Name]
EFFECTIVE DATE: [Date]

POLICY OVERVIEW
[Brief overview of what this policy covers]

GUIDELINES
[Detailed guidelines and procedures]

EXCEPTIONS
[Any exceptions to this policy]

CONTACT
For questions regarding this policy, contact: [Name / Role]
"@

$existingFile = Get-PnPFile -Url "$LibraryName/General-Department-Policy.docx" -ErrorAction SilentlyContinue
if ($existingFile) {
    Write-Host "[SKIP] General-Department-Policy.docx already exists." -ForegroundColor Yellow
} else {
    Add-PnPFile -Path $docxPath3 -Folder $LibraryName -Values @{
        Title = "General Department Policy"
        TemplateType = "Word"
        Description = "General-purpose department policy template with flexible sections. Suitable for all departments."
        Category = "General"
        IsDefault = $false
        IsActive = $true
        SortOrder = 3
    }
    Write-Host "[OK] Uploaded: General-Department-Policy.docx" -ForegroundColor Green
}

# --- Template 4: Policy Data Sheet (Excel) ---
$xlsxPath = Join-Path $tempDir "Policy-Data-Sheet.xlsx"
New-MinimalXlsx -FilePath $xlsxPath -Title "Policy Data Sheet"

$existingFile = Get-PnPFile -Url "$LibraryName/Policy-Data-Sheet.xlsx" -ErrorAction SilentlyContinue
if ($existingFile) {
    Write-Host "[SKIP] Policy-Data-Sheet.xlsx already exists." -ForegroundColor Yellow
} else {
    Add-PnPFile -Path $xlsxPath -Folder $LibraryName -Values @{
        Title = "Policy Data Sheet"
        TemplateType = "Excel"
        Description = "Excel workbook for policies with data tables and compliance checklists. Includes headers for Policy Name, Category, Status, Risk Level, Effective Date, Owner."
        Category = "General"
        IsDefault = $false
        IsActive = $true
        SortOrder = 4
    }
    Write-Host "[OK] Uploaded: Policy-Data-Sheet.xlsx" -ForegroundColor Green
}

# --- Template 5: Policy Presentation (PowerPoint) ---
$pptxPath = Join-Path $tempDir "Policy-Presentation.pptx"
New-MinimalPptx -FilePath $pptxPath -Title "Policy Awareness Presentation"

$existingFile = Get-PnPFile -Url "$LibraryName/Policy-Presentation.pptx" -ErrorAction SilentlyContinue
if ($existingFile) {
    Write-Host "[SKIP] Policy-Presentation.pptx already exists." -ForegroundColor Yellow
} else {
    Add-PnPFile -Path $pptxPath -Folder $LibraryName -Values @{
        Title = "Policy Presentation Pack"
        TemplateType = "PowerPoint"
        Description = "PowerPoint template for policy awareness presentations. Includes branded title slide with DWx Policy Manager branding."
        Category = "Corporate"
        IsDefault = $false
        IsActive = $true
        SortOrder = 5
    }
    Write-Host "[OK] Uploaded: Policy-Presentation.pptx" -ForegroundColor Green
}

# --------------------------------------------------
# 4. Clean up temp files
# --------------------------------------------------
Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

# --------------------------------------------------
# 5. Summary
# --------------------------------------------------
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " $LibraryName provisioning complete!" -ForegroundColor Green
Write-Host "============================================`n" -ForegroundColor Cyan
Write-Host "Library URL: $SiteUrl/$LibraryName" -ForegroundColor White
Write-Host ""
Write-Host "Templates uploaded:" -ForegroundColor White
Write-Host "  1. Corporate Policy - Standard A4 (Word) [DEFAULT]" -ForegroundColor White
Write-Host "  2. Corporate Policy - Executive Brief (Word)" -ForegroundColor White
Write-Host "  3. General Department Policy (Word)" -ForegroundColor White
Write-Host "  4. Policy Data Sheet (Excel)" -ForegroundColor White
Write-Host "  5. Policy Presentation Pack (PowerPoint)" -ForegroundColor White
Write-Host ""
Write-Host "To add your own branded templates:" -ForegroundColor Yellow
Write-Host "  1. Upload .docx/.xlsx/.pptx files to the library" -ForegroundColor Yellow
Write-Host "  2. Set TemplateType, Description, Category, IsActive=Yes" -ForegroundColor Yellow
Write-Host "  3. They will appear in Policy Builder > Corporate Template" -ForegroundColor Yellow
