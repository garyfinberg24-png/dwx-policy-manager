# DWx Policy Manager - Prerequisites Checklist

**Version**: 1.2.5 | **Date**: 30 March 2026 | **Company**: First Digital

Print this checklist and complete all items before beginning deployment.

---

## Microsoft 365 Tenant

- [ ] Microsoft 365 E3/E5 (or equivalent) tenant with SharePoint Online
- [ ] SharePoint Admin access confirmed for deploying user
- [ ] Tenant App Catalog exists (`/sites/AppCatalog`)
  - If not: SharePoint Admin Center > More features > Apps > App Catalog > Create
- [ ] API Management access (for approving SPFx permission requests)
  - SharePoint Admin Center > Advanced > API access

## SharePoint Site

- [ ] Site collection created: `https://<tenant>.sharepoint.com/sites/PolicyManager`
- [ ] Deploying user has Site Collection Administrator rights
- [ ] Site storage quota: minimum 1 GB (recommended 5 GB)
- [ ] Site template: Communication Site or Team Site (both work)

## Azure Subscription (for AI & Automation Features)

- [ ] Azure subscription active with billing configured
- [ ] Deploying user has Contributor role on subscription (or target resource groups)
- [ ] Azure OpenAI service available in target region (swedencentral recommended)
  - Note: Azure OpenAI requires a separate access request from Microsoft
- [ ] GPT-4o model deployment approved and available
- [ ] Budget approved for Azure resources (~$15--80/month depending on usage)

## Tools & Software

- [ ] **PnP PowerShell** module installed
  ```powershell
  Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
  ```
- [ ] **Azure CLI** installed (for Azure deployments only)
  ```bash
  az --version  # Should be 2.50+
  ```
- [ ] **Node.js** 18.17.1+ installed (for building from source only)
  ```bash
  node --version  # Should be 18.x, 20.x, or 22.x
  ```
- [ ] **Git** installed (for cloning the repository)

## Permissions & Accounts

- [ ] **SharePoint Admin** --- deploy .sppkg, create site, approve API permissions
- [ ] **Azure Contributor** --- deploy Azure Functions, Logic Apps, OpenAI
- [ ] **Application Administrator** (Entra ID) --- register app if needed
- [ ] Service account email identified for Logic App email sender (optional)
  - Recommended: shared mailbox (e.g., `noreply@company.com`)

## Network & DNS

- [ ] SharePoint Online accessible from deployment machine
- [ ] Azure Portal accessible (`portal.azure.com`)
- [ ] No firewall rules blocking Azure Function endpoints
- [ ] No conditional access policies blocking PnP PowerShell connections

## Client Data Prepared

- [ ] User list ready (see `02-Client-Data-Templates.md`)
  - FirstName, LastName, Email, Department, JobTitle, Location, Role
- [ ] Policy categories defined (or using defaults)
- [ ] Department list confirmed
- [ ] Approval workflow requirements documented
- [ ] Branding assets ready (company logo, optional custom theme colours)

## Deployment Files Available

- [ ] `policy-manager.sppkg` package (~9.0 MB)
  - Built from source: `gulp clean && gulp bundle --ship && gulp package-solution --ship`
  - Or: obtained from release artifacts
- [ ] PowerShell provisioning scripts (in `scripts/policy-management/`)
- [ ] Azure Bicep templates (in `azure-functions/*/infra/`)

---

## Sign-Off

| Item | Completed By | Date |
|------|-------------|------|
| All prerequisites confirmed | | |
| Deployment approved to proceed | | |

---

*DWx Policy Manager v1.2.5 --- First Digital*
