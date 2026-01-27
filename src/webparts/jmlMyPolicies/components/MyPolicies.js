var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import { Stack, Text, Spinner, SpinnerSize, MessageBar, MessageBarType, DefaultButton, PrimaryButton, ProgressIndicator, Icon, Label } from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyPackService } from '../../../services/PolicyPackService';
import styles from './MyPolicies.module.scss';
var MyPolicies = /** @class */ (function (_super) {
    __extends(MyPolicies, _super);
    function MyPolicies(props) {
        var _this = _super.call(this, props) || this;
        _this.handleRefresh = function () { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ refreshing: true });
                        return [4 /*yield*/, this.loadPersonalView()];
                    case 1:
                        _a.sent();
                        this.setState({ refreshing: false });
                        return [2 /*return*/];
                }
            });
        }); };
        _this.handleAcknowledge = function (policyId) {
            var url = "".concat(window.location.origin).concat(window.location.pathname, "?policyId=").concat(policyId);
            window.location.href = url;
        };
        _this.state = {
            loading: true,
            error: null,
            personalView: null,
            refreshing: false
        };
        _this.policyPackService = new PolicyPackService(props.sp);
        return _this;
    }
    MyPolicies.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        injectPortalStyles();
                        return [4 /*yield*/, this.loadPersonalView()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    MyPolicies.prototype.loadPersonalView = function () {
        return __awaiter(this, void 0, void 0, function () {
            var personalView, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        this.setState({ loading: true, error: null });
                        return [4 /*yield*/, this.policyPackService.initialize()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.policyPackService.getPersonalPolicyView()];
                    case 2:
                        personalView = _a.sent();
                        this.setState({ personalView: personalView, loading: false });
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        console.error('Failed to load personal policy view:', error_1);
                        this.setState({
                            error: 'Failed to load your policies. Please try again later.',
                            loading: false
                        });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    MyPolicies.prototype.renderComplianceScore = function () {
        var personalView = this.state.personalView;
        var showComplianceScore = this.props.showComplianceScore;
        if (!showComplianceScore || !personalView)
            return null;
        var score = personalView.complianceScore;
        var scoreColor = score >= 90 ? '#107C10' : score >= 70 ? '#FFA500' : '#D13438';
        return (React.createElement("div", { className: styles.complianceCard },
            React.createElement(Stack, { tokens: { childrenGap: 8 } },
                React.createElement(Text, { variant: "large", className: styles.cardTitle },
                    React.createElement(Icon, { iconName: "CheckMark", className: styles.titleIcon }),
                    "Compliance Score"),
                React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 16 }, verticalAlign: "center" },
                    React.createElement("div", { className: styles.scoreCircle, style: { borderColor: scoreColor } },
                        React.createElement(Text, { variant: "xxLarge", style: { color: scoreColor, fontWeight: 600 } },
                            score,
                            "%")),
                    React.createElement(Stack, { tokens: { childrenGap: 4 } },
                        React.createElement(Text, { variant: "medium" }, score >= 90 ? 'Excellent compliance!' : score >= 70 ? 'Good progress' : 'Needs attention'),
                        React.createElement(Text, { variant: "small", className: styles.subText }, "Keep up with your policy acknowledgements"))))));
    };
    MyPolicies.prototype.renderUrgentPolicies = function () {
        var _this = this;
        var personalView = this.state.personalView;
        if (!personalView || personalView.urgentPolicies.length === 0)
            return null;
        return (React.createElement("div", { className: styles.urgentSection },
            React.createElement(Stack, { tokens: { childrenGap: 12 } },
                React.createElement(Text, { variant: "xLarge", className: styles.sectionTitle },
                    React.createElement(Icon, { iconName: "Warning", className: styles.urgentIcon }),
                    "Urgent - Due in 24 Hours"),
                personalView.urgentPolicies.map(function (ack) { return (React.createElement("div", { key: ack.Id, className: styles.policyCard + ' ' + styles.urgentCard },
                    React.createElement(Stack, { tokens: { childrenGap: 8 } },
                        React.createElement(Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center" },
                            React.createElement(Text, { variant: "large", className: styles.policyTitle },
                                ack.PolicyNumber,
                                " - ",
                                ack.PolicyName),
                            React.createElement(Icon, { iconName: "Clock", className: styles.clockIcon })),
                        React.createElement(Text, { variant: "small", className: styles.category }, ack.PolicyCategory),
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 8 }, verticalAlign: "center" },
                            React.createElement(Icon, { iconName: "Calendar", className: styles.icon }),
                            React.createElement(Text, { variant: "small" },
                                "Due: ",
                                new Date(ack.DueDate).toLocaleDateString())),
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 8 }, className: styles.actions },
                            React.createElement(PrimaryButton, { text: "Acknowledge Now", iconProps: { iconName: 'Accept' }, onClick: function () { return _this.handleAcknowledge(ack.PolicyId); } }))))); }))));
    };
    MyPolicies.prototype.renderDueSoonPolicies = function () {
        var _this = this;
        var personalView = this.state.personalView;
        if (!personalView || personalView.dueSoon.length === 0)
            return null;
        return (React.createElement("div", { className: styles.dueSoonSection },
            React.createElement(Stack, { tokens: { childrenGap: 12 } },
                React.createElement(Text, { variant: "xLarge", className: styles.sectionTitle },
                    React.createElement(Icon, { iconName: "Clock", className: styles.dueSoonIcon }),
                    "Due Soon (Next 7 Days)"),
                personalView.dueSoon.map(function (ack) { return (React.createElement("div", { key: ack.Id, className: styles.policyCard },
                    React.createElement(Stack, { tokens: { childrenGap: 8 } },
                        React.createElement(Text, { variant: "large", className: styles.policyTitle },
                            ack.PolicyNumber,
                            " - ",
                            ack.PolicyName),
                        React.createElement(Text, { variant: "small", className: styles.category }, ack.PolicyCategory),
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 8 }, verticalAlign: "center" },
                            React.createElement(Icon, { iconName: "Calendar", className: styles.icon }),
                            React.createElement(Text, { variant: "small" },
                                "Due: ",
                                new Date(ack.DueDate).toLocaleDateString())),
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 8 }, className: styles.actions },
                            React.createElement(DefaultButton, { text: "View & Acknowledge", iconProps: { iconName: 'View' }, onClick: function () { return _this.handleAcknowledge(ack.PolicyId); } }))))); }))));
    };
    MyPolicies.prototype.renderPolicyPacks = function () {
        var personalView = this.state.personalView;
        var showPolicyPacks = this.props.showPolicyPacks;
        if (!showPolicyPacks || !personalView || personalView.activePolicyPacks.length === 0)
            return null;
        return (React.createElement("div", { className: styles.policyPacksSection },
            React.createElement(Stack, { tokens: { childrenGap: 12 } },
                React.createElement(Text, { variant: "xLarge", className: styles.sectionTitle },
                    React.createElement(Icon, { iconName: "BulletedList", className: styles.packIcon }),
                    "Active Policy Packs"),
                personalView.activePolicyPacks.map(function (pack) { return (React.createElement("div", { key: pack.assignmentId, className: styles.packCard },
                    React.createElement(Stack, { tokens: { childrenGap: 12 } },
                        React.createElement(Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center" },
                            React.createElement(Text, { variant: "large", className: styles.packTitle }, pack.packName),
                            React.createElement(Text, { variant: "medium", style: { fontWeight: 600 } },
                                pack.progressPercentage,
                                "%")),
                        React.createElement(ProgressIndicator, { percentComplete: pack.progressPercentage / 100, barHeight: 8 }),
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 16 } },
                            React.createElement(Stack, { tokens: { childrenGap: 4 } },
                                React.createElement(Text, { variant: "small", className: styles.packStat },
                                    pack.acknowledgements.filter(function (a) { return a.Status === 'Acknowledged'; }).length,
                                    " / ",
                                    pack.acknowledgements.length),
                                React.createElement(Text, { variant: "xSmall", className: styles.subText }, "Completed")),
                            React.createElement(Stack, { tokens: { childrenGap: 4 } },
                                React.createElement(Text, { variant: "small", className: styles.packStat }, pack.acknowledgements.filter(function (a) { return a.Status === 'Overdue'; }).length),
                                React.createElement(Text, { variant: "xSmall", className: styles.subText }, "Overdue")),
                            pack.estimatedCompletionDate && (React.createElement(Stack, { tokens: { childrenGap: 4 } },
                                React.createElement(Text, { variant: "small", className: styles.packStat }, new Date(pack.estimatedCompletionDate).toLocaleDateString()),
                                React.createElement(Text, { variant: "xSmall", className: styles.subText }, "Est. Completion")))),
                        !pack.isOnTrack && (React.createElement(MessageBar, { messageBarType: MessageBarType.warning }, "This policy pack is behind schedule. Please review urgent items."))))); }))));
    };
    MyPolicies.prototype.renderJMLIntegration = function () {
        var personalView = this.state.personalView;
        var showJMLIntegration = this.props.showJMLIntegration;
        if (!showJMLIntegration || !(personalView === null || personalView === void 0 ? void 0 : personalView.jmlIntegration))
            return null;
        var jml = personalView.jmlIntegration;
        return (React.createElement("div", { className: styles.jmlSection },
            React.createElement(Stack, { tokens: { childrenGap: 12 } },
                React.createElement(Text, { variant: "xLarge", className: styles.sectionTitle },
                    React.createElement(Icon, { iconName: "People", className: styles.jmlIcon }),
                    "Onboarding Progress"),
                React.createElement("div", { className: styles.jmlCard },
                    React.createElement(Stack, { tokens: { childrenGap: 12 } },
                        React.createElement(Stack, { horizontal: true, horizontalAlign: "space-between" },
                            React.createElement(Text, { variant: "large" },
                                jml.processType,
                                " Process"),
                            React.createElement(Text, { variant: "medium", style: { fontWeight: 600 } }, jml.processStatus)),
                        React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 8 }, verticalAlign: "center" },
                            React.createElement(Icon, { iconName: "Calendar", className: styles.icon }),
                            React.createElement(Text, { variant: "small" },
                                "Start Date: ",
                                new Date(jml.processStartDate).toLocaleDateString())),
                        jml.currentStage && (React.createElement(Stack, { tokens: { childrenGap: 8 } },
                            React.createElement(Label, null,
                                "Current Stage: ",
                                jml.currentStage),
                            jml.stageCompliance && jml.stageCompliance[jml.currentStage] && (React.createElement(Stack, { tokens: { childrenGap: 4 } },
                                React.createElement(Text, { variant: "small" },
                                    "Stage Compliance: ",
                                    jml.stageCompliance[jml.currentStage].acknowledgedCount,
                                    " / ",
                                    jml.stageCompliance[jml.currentStage].totalPolicies,
                                    " policies"),
                                React.createElement(ProgressIndicator, { percentComplete: jml.stageCompliance[jml.currentStage].acknowledgedCount /
                                        jml.stageCompliance[jml.currentStage].totalPolicies, barHeight: 6 }))))),
                        jml.blockingPolicies && jml.blockingPolicies.length > 0 && (React.createElement(MessageBar, { messageBarType: MessageBarType.severeWarning },
                            jml.blockingPolicies.length,
                            " blocking policies must be acknowledged to proceed")))))));
    };
    MyPolicies.prototype.renderEmptyState = function () {
        return (React.createElement("div", { className: styles.emptyState },
            React.createElement(Stack, { tokens: { childrenGap: 16 }, horizontalAlign: "center" },
                React.createElement(Icon, { iconName: "CompletedSolid", className: styles.emptyIcon }),
                React.createElement(Text, { variant: "xLarge" }, "All Caught Up!"),
                React.createElement(Text, { variant: "medium", className: styles.subText }, "You have no pending policy acknowledgements at this time."))));
    };
    MyPolicies.prototype.render = function () {
        var _a = this.state, loading = _a.loading, error = _a.error, personalView = _a.personalView, refreshing = _a.refreshing;
        return (React.createElement(JmlAppLayout, { context: this.props.context, pageTitle: "My Policies", pageIcon: "DocumentSet", activeNavKey: "policies", showQuickLinks: true, showSearch: true, showNotifications: true, compactFooter: true },
            React.createElement("section", { className: styles.myPolicies },
                React.createElement(Stack, { tokens: { childrenGap: 24 } },
                    React.createElement(Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center" },
                        React.createElement(DefaultButton, { text: "Refresh", iconProps: { iconName: 'Refresh' }, onClick: this.handleRefresh, disabled: loading || refreshing })),
                    loading && (React.createElement(Stack, { horizontalAlign: "center", tokens: { padding: 40 } },
                        React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading your policies..." }))),
                    error && (React.createElement(MessageBar, { messageBarType: MessageBarType.error, isMultiline: true }, error)),
                    !loading && !error && personalView && (React.createElement(Stack, { tokens: { childrenGap: 24 } },
                        this.renderComplianceScore(),
                        this.renderUrgentPolicies(),
                        this.renderDueSoonPolicies(),
                        this.renderPolicyPacks(),
                        this.renderJMLIntegration(),
                        personalView.urgentPolicies.length === 0 &&
                            personalView.dueSoon.length === 0 &&
                            personalView.activePolicyPacks.length === 0 &&
                            this.renderEmptyState()))))));
    };
    return MyPolicies;
}(React.Component));
export default MyPolicies;
//# sourceMappingURL=MyPolicies.js.map