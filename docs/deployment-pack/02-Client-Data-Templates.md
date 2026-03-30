# DWx Policy Manager - Client Data Templates

**Version**: 1.2.5 | **Date**: 30 March 2026 | **Company**: First Digital

This document contains CSV templates for data the client must provide before or during deployment. Save each section as a separate `.csv` file and populate with your organisation's data.

---

## 1. Users

**Target List**: `PM_UserProfiles`
**Required**: Yes (minimum: all employees who will use Policy Manager)

```csv
FirstName,LastName,Email,Department,JobTitle,Location,Role
John,Smith,john.smith@company.com,Human Resources,HR Director,Sydney,Manager
Jane,Doe,jane.doe@company.com,Legal,Legal Counsel,Melbourne,Author
Alice,Johnson,alice.johnson@company.com,IT & Security,CISO,Sydney,Admin
Bob,Williams,bob.williams@company.com,Finance,Financial Analyst,Brisbane,User
Carol,Brown,carol.brown@company.com,Operations,Operations Manager,Sydney,Manager
Dave,Wilson,dave.wilson@company.com,Compliance,Compliance Officer,Melbourne,Author
Eve,Taylor,eve.taylor@company.com,HR Policies,HR Coordinator,Sydney,User
Frank,Anderson,frank.anderson@company.com,IT & Security,IT Manager,Brisbane,Manager
```

### Column Definitions

| Column | Required | Description | Valid Values |
|--------|----------|-------------|-------------|
| FirstName | Yes | User's given name | Free text |
| LastName | Yes | User's surname | Free text |
| Email | Yes | Primary email address (must match Entra ID UPN) | Valid email |
| Department | Yes | Department name (must match policy audience targeting) | Free text |
| JobTitle | No | Job title | Free text |
| Location | No | Office location | Free text |
| Role | Yes | Policy Manager role assignment | `User`, `Author`, `Manager`, `Admin` |

### Role Guidance

| Role | Assign To | Typical Count |
|------|----------|---------------|
| **User** | All employees who need to read and acknowledge policies | 80--95% of users |
| **Author** | Policy writers, compliance officers, subject matter experts | 3--10% of users |
| **Manager** | Department heads, team leads who approve and track compliance | 3--8% of users |
| **Admin** | IT administrators, system owners | 1--3 people |

> **Note**: Manager does NOT inherit Author permissions. If someone needs to both write policies and manage team compliance, assign them the **Admin** role.

---

## 2. Policy Categories

**Target List**: `PM_PolicyCategories`
**Required**: No (defaults are provided, but custom categories recommended)

```csv
CategoryName,SortOrder,Description,IsActive
HR Policies,1,Human resources and employment policies,TRUE
IT & Security,2,Information technology and cybersecurity policies,TRUE
Health & Safety,3,Workplace health and safety requirements,TRUE
Compliance,4,Regulatory compliance and governance,TRUE
Financial,5,Financial management and reporting policies,TRUE
Operational,6,Business operations and process policies,TRUE
Legal,7,Legal and contractual policies,TRUE
Environmental,8,Environmental and sustainability policies,TRUE
Quality Assurance,9,Quality management standards,TRUE
Data Privacy,10,Data protection and privacy regulations,TRUE
```

### Column Definitions

| Column | Required | Description |
|--------|----------|-------------|
| CategoryName | Yes | Display name for the category |
| SortOrder | Yes | Numeric sort order (1 = first) |
| Description | No | Brief description shown in tooltips |
| IsActive | Yes | Whether the category is available for selection (`TRUE`/`FALSE`) |

---

## 3. Departments

**Target**: Used for audience targeting in policy distribution. Department names must match exactly between `PM_UserProfiles.Department` and policy audience rules.

```csv
Name,ManagerEmail,Location
Human Resources,john.smith@company.com,Sydney
IT & Security,alice.johnson@company.com,Sydney
Legal,jane.doe@company.com,Melbourne
Finance,bob.williams@company.com,Brisbane
Operations,carol.brown@company.com,Sydney
Compliance,dave.wilson@company.com,Melbourne
Marketing,sarah.jones@company.com,Sydney
Sales,mike.chen@company.com,Brisbane
Research & Development,lisa.wang@company.com,Melbourne
Customer Support,tom.garcia@company.com,Sydney
```

### Column Definitions

| Column | Required | Description |
|--------|----------|-------------|
| Name | Yes | Department name (must match PM_UserProfiles.Department exactly) |
| ManagerEmail | No | Department manager's email (for approval routing) |
| Location | No | Primary office location |

> **Important**: Department names are case-sensitive in audience targeting. "IT & Security" and "IT and Security" are treated as different departments.

---

## 4. Audiences

**Target List**: `PM_Audiences` (V2) / Audience Rules in policy distribution
**Required**: No (audiences can be created in the Admin Centre UI)

```csv
AudienceName,Category,Description,Rules
All Employees,Global,All active employees in the organisation,"{""type"":""all""}"
Sydney Office,Location,All employees based in Sydney,"{""type"":""location"",""values"":[""Sydney""]}"
IT Department,Department,Information Technology team,"{""type"":""department"",""values"":[""IT & Security""]}"
New Hires (30 days),Onboarding,Employees joined in the last 30 days,"{""type"":""newHire"",""days"":30}"
Managers Only,Role,All department managers,"{""type"":""role"",""values"":[""Manager"",""Admin""]}"
Finance & Legal,Cross-Dept,Finance and Legal departments combined,"{""type"":""department"",""values"":[""Finance"",""Legal""]}"
```

### Column Definitions

| Column | Required | Description |
|--------|----------|-------------|
| AudienceName | Yes | Display name for the audience |
| Category | No | Grouping category (Global, Department, Location, Role, Onboarding, Cross-Dept) |
| Description | No | Brief description of who this audience includes |
| Rules | Yes | JSON rule definition (see rule types below) |

### Audience Rule Types

| Type | Description | Example |
|------|-------------|---------|
| `all` | All active users | `{"type":"all"}` |
| `department` | Users in specified departments | `{"type":"department","values":["HR","Legal"]}` |
| `location` | Users at specified locations | `{"type":"location","values":["Sydney"]}` |
| `role` | Users with specified PM roles | `{"type":"role","values":["Author","Manager"]}` |
| `jobTitle` | Users with specified job titles | `{"type":"jobTitle","values":["Director","VP"]}` |
| `newHire` | Users who joined within N days | `{"type":"newHire","days":30}` |
| `securityGroup` | Members of Entra ID security groups | `{"type":"securityGroup","groupIds":["<guid>"]}` |

---

## 5. Approval Workflow Templates

**Target List**: `PM_ApprovalTemplates`
**Required**: No (can be configured in Admin Centre)

```csv
Name,Type,Description,Levels,EscalationDays,AutoApproveAfterDays
Fast Track,Single,Single approver for low-risk policies,1,3,0
Standard Review,Sequential,Two-level sequential approval,2,5,0
Regulatory Approval,Sequential,Three-level approval for regulatory policies,3,3,14
Executive Sign-Off,Parallel,Parallel approval by executive team,1,7,0
Department Head Only,Single,Department head approval only,1,5,0
```

### Column Definitions

| Column | Required | Description |
|--------|----------|-------------|
| Name | Yes | Template display name |
| Type | Yes | `Single` (one approver), `Sequential` (levels in order), `Parallel` (all at once) |
| Description | No | Brief description of when to use this template |
| Levels | Yes | Number of approval levels (1--5) |
| EscalationDays | No | Days before auto-escalation to next approver (0 = no escalation) |
| AutoApproveAfterDays | No | Days before auto-approval if no response (0 = never auto-approve) |

### Notes on Approval Configuration

- Actual approvers are assigned per policy at submission time, not in the template
- Templates define the workflow structure; the Admin Centre maps templates to policy types
- Escalation requires the Approval Escalation Logic App to be deployed (see Deployment Guide Section 3.7)

---

## Data Preparation Tips

1. **Export from Entra ID**: The easiest way to get user data is to export from Azure Active Directory > Users > Download users (CSV)
2. **Department consistency**: Ensure department names in the Users CSV match exactly what you use in Audiences and Category assignments
3. **Start with defaults**: You can deploy with default categories and configure custom ones later in the Admin Centre
4. **Minimum viable data**: At a minimum, provide the Users CSV. Everything else can be configured through the Admin Centre UI after deployment.

---

*DWx Policy Manager v1.2.5 --- First Digital*
