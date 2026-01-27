// @ts-nocheck
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { PolicyLists, TemplateLibraryLists } from '../constants/SharePointListNames';

// Template-specific list constants (extend SharePointListNames if needed for provisioning)
const TEMPLATE_LISTS = {
  TEMPLATE_VERSIONS: 'PM_PolicyTemplateVersions',
  TEMPLATE_APPROVALS: 'PM_TemplateApprovals',
  TEMPLATE_CHANGE_LOG: 'PM_TemplateChangeLog',
  VALIDATION_RULES: 'PM_PolicyValidationRules',
} as const;

export interface ITemplateVersion {
  Id: number;
  TemplateId: number;
  TemplateName: string;
  VersionNumber: string;
  VersionLabel: string;
  TemplateContent: string;
  TemplateType: string;
  TemplateCategory: string;
  ChangeDescription: string;
  ChangedBy: any;
  ChangedDate: string;
  IsCurrentVersion: boolean;
  ComplianceRisk: string;
  KeyPointsTemplate: string;
  RequiresAcknowledgement: boolean;
  RequiresQuiz: boolean;
}

export interface ITemplateApproval {
  Id: number;
  TemplateId: number;
  TemplateName: string;
  VersionNumber: string;
  ApprovalStatus: string;
  SubmittedBy: any;
  SubmittedDate: string;
  Approver: any;
  ApprovedDate?: string;
  ApprovalComments?: string;
  RejectionReason?: string;
  RequestedChanges?: string;
  ChangesSummary: string;
  ApprovalPriority: string;
  DueDate?: string;
}

export interface ITemplateChangeLog {
  Id: number;
  TemplateId: number;
  TemplateName: string;
  ActionType: string;
  VersionFrom?: string;
  VersionTo?: string;
  ActionBy: any;
  ActionDate: string;
  ActionDescription: string;
  IPAddress?: string;
}

export interface IValidationRule {
  Id: number;
  RuleName: string;
  RuleType: string;
  FieldName: string;
  RuleCondition: string;
  ErrorMessage: string;
  Severity: string;
  AppliesTo: string;
  CategoryFilter?: string;
  RiskLevelFilter?: string;
  IsActive: boolean;
  RuleOrder: number;
}

export interface IValidationResult {
  isValid: boolean;
  errors: IValidationError[];
  warnings: IValidationError[];
}

export interface IValidationError {
  field: string;
  message: string;
  severity: string;
  rule: string;
}

export class TemplateVersionService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // Version Control Methods
  // ============================================================================

  /**
   * Create a new version of a template
   */
  public async createVersion(
    templateId: number,
    templateData: any,
    changeDescription: string,
    currentUser: any
  ): Promise<ITemplateVersion> {
    try {
      // Get current template
      const template = await this.sp.web.lists
        .getByTitle("PolicyLists.POLICY_TEMPLATES")
        .items.getById(templateId)();

      // Get current version number and increment
      const currentVersion = template.CurrentVersion || "1.0";
      const versionParts = currentVersion.split(".");
      const majorVersion = parseInt(versionParts[0]);
      const minorVersion = parseInt(versionParts[1]) + 1;
      const newVersion = `${majorVersion}.${minorVersion}`;

      // Mark all previous versions as not current
      const previousVersions = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.filter(`TemplateId eq ${templateId} and IsCurrentVersion eq true`)();

      for (const prevVersion of previousVersions) {
        await this.sp.web.lists
          .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
          .items.getById(prevVersion.Id)
          .update({ IsCurrentVersion: false });
      }

      // Create new version record
      const versionData = {
        TemplateId: templateId,
        TemplateName: templateData.Title || template.Title,
        VersionNumber: newVersion,
        VersionLabel: `Version ${newVersion}`,
        TemplateContent: templateData.TemplateContent || template.TemplateContent,
        TemplateType: templateData.TemplateType || template.TemplateType,
        TemplateCategory: templateData.TemplateCategory || template.TemplateCategory,
        ChangeDescription: changeDescription,
        ChangedById: currentUser.Id,
        ChangedDate: new Date().toISOString(),
        IsCurrentVersion: true,
        ComplianceRisk: templateData.ComplianceRisk || template.ComplianceRisk,
        KeyPointsTemplate: templateData.KeyPointsTemplate || template.KeyPointsTemplate,
        RequiresAcknowledgement: templateData.RequiresAcknowledgement ?? template.RequiresAcknowledgement,
        RequiresQuiz: templateData.RequiresQuiz ?? template.RequiresQuiz
      };

      const newVersionItem = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.add(versionData);

      // Update template with new version number
      await this.sp.web.lists
        .getByTitle("PolicyLists.POLICY_TEMPLATES")
        .items.getById(templateId)
        .update({
          CurrentVersion: newVersion,
          ...templateData
        });

      // Log the version creation
      await this.logChange(
        templateId,
        templateData.Title || template.Title,
        "Version Created",
        currentVersion,
        newVersion,
        currentUser,
        `Created version ${newVersion}: ${changeDescription}`
      );

      return newVersionItem.data as ITemplateVersion;
    } catch (error) {
      console.error("Failed to create version:", error);
      throw error;
    }
  }

  /**
   * Get version history for a template
   */
  public async getVersionHistory(templateId: number): Promise<ITemplateVersion[]> {
    try {
      const versions = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.filter(`TemplateId eq ${templateId}`)
        .orderBy("VersionNumber", false)
        .top(100)
        .expand("ChangedBy")();

      return versions as ITemplateVersion[];
    } catch (error) {
      console.error("Failed to get version history:", error);
      return [];
    }
  }

  /**
   * Rollback to a previous version
   */
  public async rollbackToVersion(
    templateId: number,
    versionId: number,
    currentUser: any
  ): Promise<void> {
    try {
      // Get the target version
      const targetVersion = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.getById(versionId)();

      // Get current template
      const currentTemplate = await this.sp.web.lists
        .getByTitle("PolicyLists.POLICY_TEMPLATES")
        .items.getById(templateId)();

      const currentVersion = currentTemplate.CurrentVersion;

      // Mark all versions as not current
      const allVersions = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.filter(`TemplateId eq ${templateId}`)();

      for (const version of allVersions) {
        await this.sp.web.lists
          .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
          .items.getById(version.Id)
          .update({ IsCurrentVersion: false });
      }

      // Update the target version as current
      await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.getById(versionId)
        .update({ IsCurrentVersion: true });

      // Update the template with the rolled back content
      await this.sp.web.lists
        .getByTitle("PolicyLists.POLICY_TEMPLATES")
        .items.getById(templateId)
        .update({
          TemplateContent: targetVersion.TemplateContent,
          TemplateType: targetVersion.TemplateType,
          TemplateCategory: targetVersion.TemplateCategory,
          ComplianceRisk: targetVersion.ComplianceRisk,
          KeyPointsTemplate: targetVersion.KeyPointsTemplate,
          RequiresAcknowledgement: targetVersion.RequiresAcknowledgement,
          RequiresQuiz: targetVersion.RequiresQuiz,
          CurrentVersion: targetVersion.VersionNumber
        });

      // Log the rollback
      await this.logChange(
        templateId,
        targetVersion.TemplateName,
        "Rollback",
        currentVersion,
        targetVersion.VersionNumber,
        currentUser,
        `Rolled back from version ${currentVersion} to ${targetVersion.VersionNumber}`
      );
    } catch (error) {
      console.error("Failed to rollback version:", error);
      throw error;
    }
  }

  /**
   * Compare two versions
   */
  public async compareVersions(
    versionId1: number,
    versionId2: number
  ): Promise<{ version1: ITemplateVersion; version2: ITemplateVersion }> {
    try {
      const version1 = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.getById(versionId1)
        .expand("ChangedBy")();

      const version2 = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_VERSIONS")
        .items.getById(versionId2)
        .expand("ChangedBy")();

      return {
        version1: version1 as ITemplateVersion,
        version2: version2 as ITemplateVersion
      };
    } catch (error) {
      console.error("Failed to compare versions:", error);
      throw error;
    }
  }

  // ============================================================================
  // Approval Workflow Methods
  // ============================================================================

  /**
   * Submit template for approval
   */
  public async submitForApproval(
    templateId: number,
    templateName: string,
    versionNumber: string,
    approver: any,
    changesSummary: string,
    priority: string,
    currentUser: any
  ): Promise<ITemplateApproval> {
    try {
      const approvalData = {
        TemplateId: templateId,
        TemplateName: templateName,
        VersionNumber: versionNumber,
        ApprovalStatus: "Pending Approval",
        SubmittedById: currentUser.Id,
        SubmittedDate: new Date().toISOString(),
        ApproverId: approver.Id,
        ChangesSummary: changesSummary,
        ApprovalPriority: priority,
        DueDate: this.calculateDueDate(priority)
      };

      const approval = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_APPROVALS")
        .items.add(approvalData);

      // Update template status
      await this.sp.web.lists
        .getByTitle("PolicyLists.POLICY_TEMPLATES")
        .items.getById(templateId)
        .update({ TemplateStatus: "Pending Approval" });

      // Log the submission
      await this.logChange(
        templateId,
        templateName,
        "Submitted for Approval",
        undefined,
        versionNumber,
        currentUser,
        `Submitted version ${versionNumber} for approval`
      );

      return approval.data as ITemplateApproval;
    } catch (error) {
      console.error("Failed to submit for approval:", error);
      throw error;
    }
  }

  /**
   * Approve template
   */
  public async approveTemplate(
    approvalId: number,
    templateId: number,
    templateName: string,
    comments: string,
    currentUser: any
  ): Promise<void> {
    try {
      // Update approval record
      await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_APPROVALS")
        .items.getById(approvalId)
        .update({
          ApprovalStatus: "Approved",
          ApprovedDate: new Date().toISOString(),
          ApprovalComments: comments
        });

      // Update template status
      await this.sp.web.lists
        .getByTitle("PolicyLists.POLICY_TEMPLATES")
        .items.getById(templateId)
        .update({
          TemplateStatus: "Approved",
          ApprovedById: currentUser.Id,
          ApprovedDate: new Date().toISOString()
        });

      // Log the approval
      await this.logChange(
        templateId,
        templateName,
        "Approved",
        undefined,
        undefined,
        currentUser,
        `Template approved: ${comments || "No comments"}`
      );
    } catch (error) {
      console.error("Failed to approve template:", error);
      throw error;
    }
  }

  /**
   * Reject template
   */
  public async rejectTemplate(
    approvalId: number,
    templateId: number,
    templateName: string,
    reason: string,
    requestedChanges: string,
    currentUser: any
  ): Promise<void> {
    try {
      // Update approval record
      await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_APPROVALS")
        .items.getById(approvalId)
        .update({
          ApprovalStatus: "Rejected",
          ApprovedDate: new Date().toISOString(),
          RejectionReason: reason,
          RequestedChanges: requestedChanges
        });

      // Update template status
      await this.sp.web.lists
        .getByTitle("PolicyLists.POLICY_TEMPLATES")
        .items.getById(templateId)
        .update({ TemplateStatus: "Revision Required" });

      // Log the rejection
      await this.logChange(
        templateId,
        templateName,
        "Rejected",
        undefined,
        undefined,
        currentUser,
        `Template rejected: ${reason}`
      );
    } catch (error) {
      console.error("Failed to reject template:", error);
      throw error;
    }
  }

  /**
   * Get pending approvals
   */
  public async getPendingApprovals(userId?: number): Promise<ITemplateApproval[]> {
    try {
      let filter = "ApprovalStatus eq 'Pending Approval'";
      if (userId) {
        filter += ` and ApproverId eq ${userId}`;
      }

      const approvals = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_APPROVALS")
        .items.filter(filter)
        .orderBy("ApprovalPriority", true)
        .orderBy("SubmittedDate", true)
        .top(100)
        .expand("SubmittedBy", "Approver")();

      return approvals as ITemplateApproval[];
    } catch (error) {
      console.error("Failed to get pending approvals:", error);
      return [];
    }
  }

  // ============================================================================
  // Validation Methods
  // ============================================================================

  /**
   * Get validation rules
   */
  public async getValidationRules(
    category?: string,
    riskLevel?: string
  ): Promise<IValidationRule[]> {
    try {
      let filter = "IsActive eq true";

      if (category) {
        filter += ` and (AppliesTo eq 'All Policies' or AppliesTo eq 'Specific Category' and CategoryFilter eq '${category}')`;
      }

      if (riskLevel) {
        filter += ` and (AppliesTo eq 'All Policies' or AppliesTo eq 'Specific Risk Level' and RiskLevelFilter eq '${riskLevel}')`;
      }

      const rules = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.VALIDATION_RULES")
        .items.filter(filter)
        .orderBy("RuleOrder", true)
        .top(100)();

      return rules as IValidationRule[];
    } catch (error) {
      console.error("Failed to get validation rules:", error);
      return [];
    }
  }

  /**
   * Validate policy data against rules
   */
  public async validatePolicy(policyData: any): Promise<IValidationResult> {
    const errors: IValidationError[] = [];
    const warnings: IValidationError[] = [];

    try {
      // Get applicable rules
      const rules = await this.getValidationRules(
        policyData.policyCategory,
        policyData.complianceRisk
      );

      for (const rule of rules) {
        const validationError = this.applyValidationRule(rule, policyData);
        if (validationError) {
          if (rule.Severity === "Error") {
            errors.push(validationError);
          } else if (rule.Severity === "Warning") {
            warnings.push(validationError);
          }
        }
      }
    } catch (error) {
      console.error("Validation failed:", error);
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Apply a single validation rule
   */
  private applyValidationRule(rule: IValidationRule, policyData: any): IValidationError | null {
    const fieldValue = policyData[this.convertFieldName(rule.FieldName)];

    switch (rule.RuleType) {
      case "Required Field":
        if (!fieldValue || fieldValue.toString().trim() === "") {
          return {
            field: rule.FieldName,
            message: rule.ErrorMessage,
            severity: rule.Severity,
            rule: rule.RuleName
          };
        }
        break;

      case "Business Rule":
        // High risk must require quiz
        if (rule.RuleName === "High Risk Policy Must Require Quiz") {
          if (policyData.complianceRisk === "High" && !policyData.requiresQuiz) {
            return {
              field: "RequiresQuiz",
              message: rule.ErrorMessage,
              severity: rule.Severity,
              rule: rule.RuleName
            };
          }
        }

        // High risk must require acknowledgement
        if (rule.RuleName === "High Risk Policy Must Require Acknowledgement") {
          if (policyData.complianceRisk === "High" && !policyData.requiresAcknowledgement) {
            return {
              field: "RequiresAcknowledgement",
              message: rule.ErrorMessage,
              severity: rule.Severity,
              rule: rule.RuleName
            };
          }
        }

        // Approver required for high risk
        if (rule.RuleName === "Approver Required for High Risk") {
          if (policyData.complianceRisk === "High" && (!policyData.approvers || policyData.approvers.length === 0)) {
            return {
              field: "Approvers",
              message: rule.ErrorMessage,
              severity: rule.Severity,
              rule: rule.RuleName
            };
          }
        }

        // Category-specific reviewer rules
        if (rule.RuleName.includes("Must Have Reviewer") || rule.RuleName.includes("Must Have Finance Reviewer")) {
          if (policyData.policyCategory === rule.CategoryFilter && (!policyData.reviewers || policyData.reviewers.length === 0)) {
            return {
              field: "Reviewers",
              message: rule.ErrorMessage,
              severity: rule.Severity,
              rule: rule.RuleName
            };
          }
        }
        break;

      case "Data Range":
        if (rule.RuleName === "Policy Title Length") {
          const title = policyData.policyTitle || "";
          if (title.length < 10 || title.length > 100) {
            return {
              field: "PolicyTitle",
              message: rule.ErrorMessage,
              severity: rule.Severity,
              rule: rule.RuleName
            };
          }
        }

        if (rule.RuleName === "Policy Content Required") {
          const content = policyData.policyContent || "";
          if (content.length < 100) {
            return {
              field: "PolicyContent",
              message: rule.ErrorMessage,
              severity: rule.Severity,
              rule: rule.RuleName
            };
          }
        }
        break;

      case "Conditional Required":
        if (rule.RuleName === "Key Points Required for New Hires") {
          const immediateTimeframes = ["Immediate", "Day 1", "Day 3"];
          if (immediateTimeframes.includes(policyData.readTimeframe)) {
            if (!policyData.keyPoints || policyData.keyPoints.length < 3) {
              return {
                field: "KeyPoints",
                message: rule.ErrorMessage,
                severity: rule.Severity,
                rule: rule.RuleName
              };
            }
          }
        }
        break;

      case "Format Validation":
        if (rule.RuleName === "Version Number Format") {
          const versionPattern = /^\d+\.\d+$/;
          if (fieldValue && !versionPattern.test(fieldValue)) {
            return {
              field: "CurrentVersion",
              message: rule.ErrorMessage,
              severity: rule.Severity,
              rule: rule.RuleName
            };
          }
        }
        break;
    }

    return null;
  }

  /**
   * Convert internal field names to data object property names
   */
  private convertFieldName(internalName: string): string {
    const fieldMap: { [key: string]: string } = {
      PolicyTitle: "policyTitle",
      PolicyCategory: "policyCategory",
      ComplianceRisk: "complianceRisk",
      RequiresQuiz: "requiresQuiz",
      RequiresAcknowledgement: "requiresAcknowledgement",
      PolicyContent: "policyContent",
      ReadTimeframe: "readTimeframe",
      Tags: "tags",
      Approvers: "approvers",
      Reviewers: "reviewers",
      KeyPoints: "keyPoints",
      CurrentVersion: "currentVersion"
    };

    return fieldMap[internalName] || internalName;
  }

  // ============================================================================
  // Change Log Methods
  // ============================================================================

  /**
   * Log a change action
   */
  private async logChange(
    templateId: number,
    templateName: string,
    actionType: string,
    versionFrom: string | undefined,
    versionTo: string | undefined,
    actionBy: any,
    description: string
  ): Promise<void> {
    try {
      const logData = {
        TemplateId: templateId,
        TemplateName: templateName,
        ActionType: actionType,
        VersionFrom: versionFrom,
        VersionTo: versionTo,
        ActionById: actionBy.Id,
        ActionDate: new Date().toISOString(),
        ActionDescription: description
      };

      await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_CHANGE_LOG")
        .items.add(logData);
    } catch (error) {
      console.error("Failed to log change:", error);
    }
  }

  /**
   * Get change log for a template
   */
  public async getChangeLog(templateId: number): Promise<ITemplateChangeLog[]> {
    try {
      const logs = await this.sp.web.lists
        .getByTitle("TEMPLATE_LISTS.TEMPLATE_CHANGE_LOG")
        .items.filter(`TemplateId eq ${templateId}`)
        .orderBy("ActionDate", false)
        .top(100)
        .expand("ActionBy")();

      return logs as ITemplateChangeLog[];
    } catch (error) {
      console.error("Failed to get change log:", error);
      return [];
    }
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  /**
   * Calculate due date based on priority
   */
  private calculateDueDate(priority: string): string {
    const now = new Date();
    let daysToAdd = 5; // Default

    switch (priority) {
      case "High":
        daysToAdd = 2;
        break;
      case "Medium":
        daysToAdd = 5;
        break;
      case "Low":
        daysToAdd = 10;
        break;
    }

    now.setDate(now.getDate() + daysToAdd);
    return now.toISOString();
  }
}
