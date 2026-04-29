using System.ComponentModel;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Linq;
using System.Diagnostics;
using NewUserAutomation.App.Services;
using NewUserAutomation.Core.Models;
using NewUserAutomation.Core.Execution;
using NewUserAutomation.Core.Parsing;
using NewUserAutomation.Core.Security;
using NewUserAutomation.Core.Validation;

namespace NewUserAutomation.App.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly FormParser _parser = new();
    private readonly ValidationEngine _validation = new();
    private readonly DryRunPlanner _dryRunPlanner = new();
    private readonly PermissionMatrix _permissionMatrix = new();
    private readonly AuthSessionManager _authSession = new();
    private readonly LiveRunLogger _liveRunLogger = new();
    private readonly CustomerProfileStore _profiles = new();
    private readonly CustomerDirectoryStore _customerDirectory = new();

    private WizardStep _currentStep = WizardStep.SignIn;
    private string _selectedFilePath = string.Empty;
    private string _statusMessage = "One-app mode: Select customer, run setup once, save, then Connect All.";
    private string _signInPhaseMessage = "Idle";
    private string _preflightResult = "Preflight not run.";
    private List<string> _dryRunLines = [];
    private NewUserRequest? _currentRequest;
    private string _editFirstName = string.Empty;
    private string _editLastName = string.Empty;
    private string _editDisplayName = string.Empty;
    private string _editPreferredUsername = string.Empty;
    private string _editTemporaryPassword = string.Empty;
    private string _editJobTitle = string.Empty;
    private string _editPrimaryEmail = string.Empty;
    private string _editSecondaryEmail = string.Empty;
    private string _editLicenseSkus = string.Empty;
    private string _editGroupAccess = string.Empty;
    private string _editSharedMailboxAccess = string.Empty;
    private string _editSharePointAccess = string.Empty;
    private string _editPersonalSharePointAdditionalMembers = string.Empty;
    private string _editSpecialRequirements = string.Empty;
    private string _editRequestApprovedBy = string.Empty;
    private bool _isAuthBusy;
    private bool _isGraphConnected;
    private bool _isExchangeConnected;
    private bool _isPnPConnected;
    private string _graphTenantId = string.Empty;
    private string _graphStatus = "Not connected";
    private string _exchangeStatus = "Not connected";
    private string _exchangeSetupStatus = "Use 'Set Up Exchange Cert' to generate/export a certificate for app registration.";
    private string _exchangeSetupThumbprint = string.Empty;
    private string _pnpStatus = "Not connected";
    private string _setupStatus = "Use Set Up Customer App (one-time) to create/reuse app registration and consent.";
    private string _selectedSharePointSiteUrl = "https://hresourcing.sharepoint.com";
    private string _pnpClientId = string.Empty;
    private string _pnpThumbprint = string.Empty;
    private string _exchangeClientId = string.Empty;
    private string _exchangeThumbprint = string.Empty;
    private string _tenantDomain = string.Empty;
    private bool _skipSharePointForNow = false;
    private CustomerDirectoryEntry? _selectedCustomer;
    private string _customerName = string.Empty;
    private bool _isCustomerEditLocked = true;
    private bool _showParserDiagnosticsDetails;
    private LiveRunState? _liveRunState;
    private string _liveRunStatus = "Live run not started.";
    private string _liveRunLogPath = string.Empty;
    private string _liveRunProgressText = "Progress: 0/0";
    private int _liveRunProgressPercent;
    private readonly string _appVersion = ResolveAppVersion();

    public event PropertyChangedEventHandler? PropertyChanged;

    public MainViewModel()
    {
        ConnectGraphCommand = new RelayCommand(ConnectGraph, () => !IsAuthBusy);
        QuickConnectCommand = new RelayCommand(QuickConnect, () => !IsAuthBusy);
        ConnectExchangeCommand = new RelayCommand(ConnectExchange, () => !IsAuthBusy);
        ShowExchangeSetupHelpCommand = new RelayCommand(ShowExchangeSetupHelp);
        SetupExchangeCertCommand = new RelayCommand(SetupExchangeCert, () => !IsAuthBusy);
        SetupCustomerAppCommand = new RelayCommand(SetupCustomerApp, () => !IsAuthBusy && !string.IsNullOrWhiteSpace(CustomerName) && !string.IsNullOrWhiteSpace(TenantDomain));
        AuthReadinessCheckCommand = new RelayCommand(RunAuthReadinessCheck, () => !IsAuthBusy);
        OpenAdminGrantScriptCommand = new RelayCommand(OpenAdminGrantScript, () => !IsAuthBusy && IsUsableSharePointUrl(SelectedSharePointSiteUrl) && !string.IsNullOrWhiteSpace(string.IsNullOrWhiteSpace(ExchangeClientId) ? PnPClientId : ExchangeClientId));
        CopyExchangeThumbprintCommand = new RelayCommand(CopyExchangeThumbprint, () => !string.IsNullOrWhiteSpace(ExchangeSetupThumbprint));
        ConnectPnPCommand = new RelayCommand(ConnectPnP, () => !IsAuthBusy && IsGraphConnected && !string.IsNullOrWhiteSpace(PnPClientId) && !string.IsNullOrWhiteSpace(PnPThumbprint) && !string.IsNullOrWhiteSpace(TenantDomain) && !string.IsNullOrWhiteSpace(SelectedSharePointSiteUrl));
        SetupPnPAppCommand = new RelayCommand(SetupPnPApp, () => !IsAuthBusy && IsGraphConnected);
        SaveCustomerCommand = new RelayCommand(
            SaveCustomer,
            () => !string.IsNullOrWhiteSpace(CustomerName)
                && !string.IsNullOrWhiteSpace(TenantDomain)
                && IsUsableSharePointUrl(SelectedSharePointSiteUrl));
        NewCustomerCommand = new RelayCommand(NewCustomer);
        DeleteCustomerCommand = new RelayCommand(DeleteCustomer, () => SelectedCustomer is not null && !IsAuthBusy);
        EditCustomerCommand = new RelayCommand(EditCustomer, () => SelectedCustomer is not null && IsCustomerEditLocked);
        ContinueToUploadCommand = new RelayCommand(() => CurrentStep = WizardStep.Upload, () => IsSignInComplete);
        ContinueToReviewCommand = new RelayCommand(ParseAndReview, () => !string.IsNullOrWhiteSpace(SelectedFilePath) && File.Exists(SelectedFilePath));
        BackToUploadCommand = new RelayCommand(() => CurrentStep = WizardStep.Upload);
        ContinueToLiveRunCommand = new RelayCommand(() => CurrentStep = WizardStep.LiveRun, () => CurrentStep == WizardStep.Review && PreflightChecks.Count > 0 && PreflightChecks.All(c => c.IsPass));
        BackToReviewCommand = new RelayCommand(() => CurrentStep = WizardStep.Review, () => CurrentStep == WizardStep.LiveRun);
        FinishLiveRunCommand = new RelayCommand(FinishLiveRun, () => CurrentStep == WizardStep.LiveRun && _liveRunState is not null && _liveRunState.Steps.Count > 0 && _liveRunState.Steps.All(step => step.IsCompleted));
        BackToLiveRunCommand = new RelayCommand(() => CurrentStep = WizardStep.LiveRun, () => CurrentStep == WizardStep.Summary);
        BackToStartCommand = new RelayCommand(BackToStart, () => !IsAuthBusy);
        CloseApplicationCommand = new RelayCommand(CloseApplication);
        RunPreflightCommand = new RelayCommand(RunPreflight, () => CurrentStep == WizardStep.Review && DryRunLines.Count > 0);
        ApplyReviewEditsCommand = new RelayCommand(ApplyReviewEdits, () => CurrentStep == WizardStep.Review && _currentRequest is not null);
        ToggleParserDiagnosticsCommand = new RelayCommand(ToggleParserDiagnostics);
        StartSafeLiveRunCommand = new RelayCommand(StartSafeLiveRun, () => CurrentStep == WizardStep.LiveRun && _currentRequest is not null);
        ExecuteNextLiveStepCommand = new RelayCommand(ExecuteNextLiveStep, () => CurrentStep == WizardStep.LiveRun && !IsAuthBusy && LiveRunSteps.Any(step => !step.IsCompleted));
        ExecuteAllLiveStepsCommand = new RelayCommand(ExecuteAllLiveSteps, () => CurrentStep == WizardStep.LiveRun && !IsAuthBusy && LiveRunSteps.Any(step => !step.IsCompleted));
        SkipLiveStepCommand = new RelayCommand(SkipLiveStep, () => CurrentStep == WizardStep.LiveRun && !IsAuthBusy && LiveRunSteps.Any(step => !step.IsCompleted));
        ResetLiveRunCommand = new RelayCommand(ResetLiveRun, () => CurrentStep == WizardStep.LiveRun && _currentRequest is not null);

        LoadCustomers();
    }

    public RelayCommand ConnectGraphCommand { get; }
    public RelayCommand QuickConnectCommand { get; }
    public RelayCommand ConnectExchangeCommand { get; }
    public RelayCommand ShowExchangeSetupHelpCommand { get; }
    public RelayCommand SetupExchangeCertCommand { get; }
    public RelayCommand SetupCustomerAppCommand { get; }
    public RelayCommand AuthReadinessCheckCommand { get; }
    public RelayCommand OpenAdminGrantScriptCommand { get; }
    public RelayCommand CopyExchangeThumbprintCommand { get; }
    public RelayCommand ConnectPnPCommand { get; }
    public RelayCommand SetupPnPAppCommand { get; }
    public RelayCommand SaveCustomerCommand { get; }
    public RelayCommand NewCustomerCommand { get; }
    public RelayCommand DeleteCustomerCommand { get; }
    public RelayCommand EditCustomerCommand { get; }
    public RelayCommand ContinueToUploadCommand { get; }
    public RelayCommand ContinueToReviewCommand { get; }
    public RelayCommand BackToUploadCommand { get; }
    public RelayCommand ContinueToLiveRunCommand { get; }
    public RelayCommand BackToReviewCommand { get; }
    public RelayCommand FinishLiveRunCommand { get; }
    public RelayCommand BackToLiveRunCommand { get; }
    public RelayCommand BackToStartCommand { get; }
    public RelayCommand CloseApplicationCommand { get; }
    public RelayCommand RunPreflightCommand { get; }
    public RelayCommand ApplyReviewEditsCommand { get; }
    public RelayCommand ToggleParserDiagnosticsCommand { get; }
    public RelayCommand StartSafeLiveRunCommand { get; }
    public RelayCommand ExecuteNextLiveStepCommand { get; }
    public RelayCommand ExecuteAllLiveStepsCommand { get; }
    public RelayCommand SkipLiveStepCommand { get; }
    public RelayCommand ResetLiveRunCommand { get; }

    public WizardStep CurrentStep
    {
        get => _currentStep;
        set
        {
            if (Set(ref _currentStep, value))
            {
                OnPropertyChanged(nameof(IsSignInStep));
                OnPropertyChanged(nameof(IsUploadStep));
                OnPropertyChanged(nameof(IsReviewStep));
                OnPropertyChanged(nameof(IsLiveRunStep));
                OnPropertyChanged(nameof(IsSummaryStep));
                RefreshCommandStates();
            }
        }
    }

    public bool IsSignInStep => CurrentStep == WizardStep.SignIn;
    public bool IsUploadStep => CurrentStep == WizardStep.Upload;
    public bool IsReviewStep => CurrentStep == WizardStep.Review;
    public bool IsLiveRunStep => CurrentStep == WizardStep.LiveRun;
    public bool IsSummaryStep => CurrentStep == WizardStep.Summary;
    public string SelectedFilePath
    {
        get => _selectedFilePath;
        set
        {
            if (Set(ref _selectedFilePath, value))
            {
                RefreshCommandStates();
            }
        }
    }

    public bool IsAuthBusy
    {
        get => _isAuthBusy;
        private set
        {
            if (Set(ref _isAuthBusy, value))
            {
                RefreshCommandStates();
            }
        }
    }

    public bool IsSignInComplete => SelectedCustomer is not null && IsGraphConnected && IsExchangeConnected && (IsPnPConnected || SkipSharePointForNow);

    public bool IsGraphConnected
    {
        get => _isGraphConnected;
        private set
        {
            if (Set(ref _isGraphConnected, value))
            {
                OnPropertyChanged(nameof(IsSignInComplete));
                RefreshCommandStates();
            }
        }
    }

    public bool IsExchangeConnected
    {
        get => _isExchangeConnected;
        private set
        {
            if (Set(ref _isExchangeConnected, value))
            {
                OnPropertyChanged(nameof(IsSignInComplete));
                RefreshCommandStates();
            }
        }
    }

    public bool IsPnPConnected
    {
        get => _isPnPConnected;
        private set
        {
            if (Set(ref _isPnPConnected, value))
            {
                OnPropertyChanged(nameof(IsSignInComplete));
                RefreshCommandStates();
            }
        }
    }

    public string GraphStatus
    {
        get => _graphStatus;
        private set => Set(ref _graphStatus, value);
    }

    public string ExchangeStatus
    {
        get => _exchangeStatus;
        private set => Set(ref _exchangeStatus, value);
    }

    public string ExchangeSetupStatus
    {
        get => _exchangeSetupStatus;
        private set => Set(ref _exchangeSetupStatus, value);
    }

    public string ExchangeSetupThumbprint
    {
        get => _exchangeSetupThumbprint;
        private set => Set(ref _exchangeSetupThumbprint, value);
    }

    public string PnPStatus
    {
        get => _pnpStatus;
        private set => Set(ref _pnpStatus, value);
    }

    public string SetupStatus
    {
        get => _setupStatus;
        private set => Set(ref _setupStatus, value);
    }

    public bool SkipSharePointForNow
    {
        get => _skipSharePointForNow;
        set
        {
            if (Set(ref _skipSharePointForNow, value))
            {
                OnPropertyChanged(nameof(IsSignInComplete));
                RefreshCommandStates();
            }
        }
    }

    public string SelectedSharePointSiteUrl
    {
        get => _selectedSharePointSiteUrl;
        set => Set(ref _selectedSharePointSiteUrl, value);
    }

    public string PnPClientId
    {
        get => _pnpClientId;
        set => Set(ref _pnpClientId, value);
    }

    public string PnPThumbprint
    {
        get => _pnpThumbprint;
        set => Set(ref _pnpThumbprint, value);
    }

    public string ExchangeClientId
    {
        get => _exchangeClientId;
        set => Set(ref _exchangeClientId, value);
    }

    public string ExchangeThumbprint
    {
        get => _exchangeThumbprint;
        set => Set(ref _exchangeThumbprint, value);
    }

    public string TenantDomain
    {
        get => _tenantDomain;
        set => Set(ref _tenantDomain, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => Set(ref _statusMessage, value);
    }

    public string SignInPhaseMessage
    {
        get => _signInPhaseMessage;
        private set => Set(ref _signInPhaseMessage, value);
    }

    public string PreflightResult
    {
        get => _preflightResult;
        private set => Set(ref _preflightResult, value);
    }

    public IReadOnlyList<string> DryRunLines => _dryRunLines;
    public string EditFirstName { get => _editFirstName; set => Set(ref _editFirstName, value); }
    public string EditLastName { get => _editLastName; set => Set(ref _editLastName, value); }
    public string EditDisplayName { get => _editDisplayName; set => Set(ref _editDisplayName, value); }
    public string EditPreferredUsername { get => _editPreferredUsername; set => Set(ref _editPreferredUsername, value); }
    public string EditTemporaryPassword { get => _editTemporaryPassword; set => Set(ref _editTemporaryPassword, value); }
    public string EditJobTitle { get => _editJobTitle; set => Set(ref _editJobTitle, value); }
    public string EditPrimaryEmail { get => _editPrimaryEmail; set => Set(ref _editPrimaryEmail, value); }
    public string EditSecondaryEmail { get => _editSecondaryEmail; set => Set(ref _editSecondaryEmail, value); }
    public string EditLicenseSkus { get => _editLicenseSkus; set => Set(ref _editLicenseSkus, value); }
    public string EditGroupAccess { get => _editGroupAccess; set => Set(ref _editGroupAccess, value); }
    public string EditSharedMailboxAccess { get => _editSharedMailboxAccess; set => Set(ref _editSharedMailboxAccess, value); }
    public string EditSharePointAccess { get => _editSharePointAccess; set => Set(ref _editSharePointAccess, value); }
    public string EditPersonalSharePointAdditionalMembers { get => _editPersonalSharePointAdditionalMembers; set => Set(ref _editPersonalSharePointAdditionalMembers, value); }
    public string EditSpecialRequirements { get => _editSpecialRequirements; set => Set(ref _editSpecialRequirements, value); }
    public string EditRequestApprovedBy { get => _editRequestApprovedBy; set => Set(ref _editRequestApprovedBy, value); }
    public ObservableCollection<PreflightCheckItem> PreflightChecks { get; } = [];
    public ObservableCollection<DirectoryUserMatchChoice> PersonalFolderUserMatchChoices { get; } = [];
    public ObservableCollection<LiveRunStepState> LiveRunSteps { get; } = [];
    public ObservableCollection<LiveRunStepState> LiveRunSummarySteps { get; } = [];
    public string LiveRunStatus
    {
        get => _liveRunStatus;
        private set => Set(ref _liveRunStatus, value);
    }
    public string LiveRunLogPath
    {
        get => _liveRunLogPath;
        private set => Set(ref _liveRunLogPath, value);
    }
    public string LiveRunSummaryHeadline
    {
        get
        {
            var skipped = LiveRunSummarySteps.Count(step => string.Equals(step.Status, "Skipped", StringComparison.OrdinalIgnoreCase));
            var completed = LiveRunSummarySteps.Count(step => string.Equals(step.Status, "Completed", StringComparison.OrdinalIgnoreCase));
            return $"Live run finished: {completed} completed, {skipped} skipped.";
        }
    }

    public string LiveRunProgressText
    {
        get => _liveRunProgressText;
        private set => Set(ref _liveRunProgressText, value);
    }

    public int LiveRunProgressPercent
    {
        get => _liveRunProgressPercent;
        private set => Set(ref _liveRunProgressPercent, value);
    }
    public string AppVersionLabel => $"App version: {_appVersion}";
    public bool ShowParserDiagnosticsDetails
    {
        get => _showParserDiagnosticsDetails;
        set
        {
            if (Set(ref _showParserDiagnosticsDetails, value))
            {
                OnPropertyChanged(nameof(ParserDiagnosticsToggleText));
            }
        }
    }
    public string ParserDiagnosticsToggleText => ShowParserDiagnosticsDetails ? "Hide Parser Diagnostics" : "Show Parser Diagnostics";

    public ObservableCollection<CustomerDirectoryEntry> Customers { get; } = [];
    public bool IsCustomerEditLocked
    {
        get => _isCustomerEditLocked;
        set
        {
            if (Set(ref _isCustomerEditLocked, value))
            {
                OnPropertyChanged(nameof(CustomerEditStatus));
                OnPropertyChanged(nameof(CustomerEditButtonText));
            }
        }
    }

    public string CustomerEditStatus => IsCustomerEditLocked ? "Viewing saved profile" : "Editing customer profile";
    public string CustomerEditButtonText => "Edit Customer";

    public CustomerDirectoryEntry? SelectedCustomer
    {
        get => _selectedCustomer;
        set
        {
            if (Set(ref _selectedCustomer, value) && value is not null)
            {
                TenantDomain = value.TenantDomain;
                SelectedSharePointSiteUrl = value.SiteUrl;
                PnPClientId = value.PnPClientId;
                PnPThumbprint = value.PnPThumbprint;
                CustomerName = value.Name;
                LoadCustomerProfileIfAvailable();
                IsCustomerEditLocked = true;
                StatusMessage = $"Selected customer: {value.Name}";
                OnPropertyChanged(nameof(IsSignInComplete));
                RefreshCommandStates();
            }
        }
    }

    public string CustomerName
    {
        get => _customerName;
        set => Set(ref _customerName, value);
    }

    public void SetSelectedFilePath(string path)
    {
        SelectedFilePath = path;
    }

    private void ConnectGraph()
    {
        _ = ConnectGraphAsync();
    }

    public void QuickConnect()
    {
        _ = QuickConnectAsync();
    }

    private void ConnectExchange()
    {
        _ = ConnectExchangeAsync();
    }

    private void ConnectPnP()
    {
        _ = ConnectPnPAsync();
    }

    private void SetupPnPApp() => _ = SetupPnPAppAsync();

    private async Task ConnectGraphAsync()
    {
        if (IsAuthBusy)
        {
            return;
        }

        IsAuthBusy = true;
        SignInPhaseMessage = "Connecting Graph...";
        try
        {
            var result = await _authSession.ConnectGraphAsync(
                appId: string.IsNullOrWhiteSpace(ExchangeClientId) ? PnPClientId : ExchangeClientId,
                tenantDomain: TenantDomain,
                thumbprint: string.IsNullOrWhiteSpace(ExchangeThumbprint) ? PnPThumbprint : ExchangeThumbprint,
                progress: new Progress<string>(phase => SignInPhaseMessage = phase));
            if (!result.Success)
            {
                var friendly = Compact(result.ErrorMessage);
                StatusMessage = "Graph sign-in failed.";
                GraphStatus = $"Failed: {friendly}";
                IsGraphConnected = false;
                return;
            }

            GraphStatus = $"Connected: {result.Account} ({result.TenantId})";
            IsGraphConnected = true;
            _graphTenantId = result.TenantId;
            if (string.IsNullOrWhiteSpace(TenantDomain) && result.Account.Contains('@'))
            {
                var domain = result.Account.Split('@').LastOrDefault();
                if (!string.IsNullOrWhiteSpace(domain))
                {
                    TenantDomain = domain;
                    LoadCustomerProfileIfAvailable();
                }
            }
            StatusMessage = "Graph connected.";
            SignInPhaseMessage = "Graph connected";
        }
        finally
        {
            IsAuthBusy = false;
        }
    }

    private async Task QuickConnectAsync()
    {
        if (IsAuthBusy) return;
        StatusMessage = "Connecting Graph, Exchange, and SharePoint with one app registration...";
        await ConnectGraphAsync();
        if (!IsGraphConnected) return;
        await ConnectExchangeAsync();
        if (!IsExchangeConnected) return;
        await ConnectPnPAsync();
        if (!IsPnPConnected) return;
        StatusMessage = "All services connected using one app registration.";
        SignInPhaseMessage = "Connect all complete";
    }

    private async Task ConnectExchangeAsync()
    {
        if (IsAuthBusy) return;
        IsAuthBusy = true;
        SignInPhaseMessage = "Connecting Exchange Online...";
        try
        {
            var result = await _authSession.ConnectExchangeAsync(
                appId: string.IsNullOrWhiteSpace(ExchangeClientId) ? PnPClientId : ExchangeClientId,
                organization: TenantDomain,
                thumbprint: string.IsNullOrWhiteSpace(ExchangeThumbprint) ? PnPThumbprint : ExchangeThumbprint,
                progress: new Progress<string>(phase => SignInPhaseMessage = phase));
            if (!result.Success)
            {
                var friendly = Compact(result.ErrorMessage);
                ExchangeStatus = $"Failed: {friendly}";
                IsExchangeConnected = false;
                StatusMessage = "Exchange sign-in failed. Re-run Quick Connect or connect Exchange directly.";
                return;
            }

            ExchangeStatus = $"Connected: {result.Account}";
            IsExchangeConnected = true;
            StatusMessage = "Exchange Online connected.";
            SignInPhaseMessage = "Exchange connected";
        }
        catch (Exception ex)
        {
            var friendly = Compact(ex.Message);
            ExchangeStatus = $"Failed: {friendly}";
            IsExchangeConnected = false;
            StatusMessage = $"Exchange sign-in failed: {friendly}";
            SignInPhaseMessage = "Exchange connection failed";
        }
        finally
        {
            IsAuthBusy = false;
        }
    }

    private void ShowExchangeSetupHelp()
    {
        var help = """
One-app setup (recommended)

Do this on the Sign In page in this exact order:

1) Enter Customer name + Tenant domain.
2) Click "Set Up Customer App (One-Time)".
   - App generates cert, configures the app registration, and opens admin consent.
3) Complete admin consent in browser.
4) Click "Save Customer".
5) Click "Connect All (One App)".

Expected result:
- Graph = Connected
- Exchange = Connected
- SharePoint = Connected

Cert files are saved to:
- artifacts/customers/<CustomerName>/exchange

If connect fails:
- Confirm admin consent completed successfully.
- Confirm Exchange.ManageAsApp and required Graph/SharePoint permissions exist.
- Confirm the cert private key is present in CurrentUser\\My on this machine.

Full guide:
- docs/exchange-app-auth-setup.md
""";
        MessageBox.Show(help, "Exchange Setup Help", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private async Task SetupExchangeCertAsync()
    {
        if (IsAuthBusy) return;
        IsAuthBusy = true;
        SignInPhaseMessage = "Generating Exchange certificate...";
        try
        {
            var safeCustomer = string.IsNullOrWhiteSpace(CustomerName)
                ? "default-customer"
                : Regex.Replace(CustomerName.Trim(), @"[^a-zA-Z0-9\-_\. ]", "_").Trim();
            if (string.IsNullOrWhiteSpace(safeCustomer))
            {
                safeCustomer = "default-customer";
            }
            var exportDir = Path.Combine(
                Directory.GetCurrentDirectory(),
                "artifacts",
                "customers",
                safeCustomer,
                "exchange");

            var result = await _authSession.EnsureExchangeCertificateAsync(
                subjectName: "NewUserAutomation.Exchange",
                exportDirectory: exportDir,
                progress: new Progress<string>(phase => SignInPhaseMessage = phase));
            if (!result.Success)
            {
                ExchangeSetupStatus = $"Exchange cert setup failed: {Compact(result.ErrorMessage)}";
                ExchangeSetupThumbprint = string.Empty;
                StatusMessage = ExchangeSetupStatus;
                return;
            }

            ExchangeSetupThumbprint = result.Thumbprint;
            ExchangeThumbprint = result.Thumbprint;
            ExchangeSetupStatus =
                $"Exchange cert ready. Subject: {result.Subject}. Exchange thumbprint: {result.Thumbprint}. Upload CER from: {result.CerPath}";
            StatusMessage =
                $"Exchange cert generated. In Entra app registration, upload CER at '{result.CerPath}', then add Exchange.ManageAsApp and grant admin consent.";
            SignInPhaseMessage = "Exchange certificate generated";
        }
        finally
        {
            IsAuthBusy = false;
        }
    }

    private void SetupExchangeCert()
    {
        _ = SetupExchangeCertAsync();
    }

    private async Task SetupCustomerAppAsync()
    {
        if (IsAuthBusy) return;
        IsAuthBusy = true;
        SignInPhaseMessage = "[RUNNING] Setting up one-time customer app...";
        try
        {
            var safeCustomer = string.IsNullOrWhiteSpace(CustomerName)
                ? "default-customer"
                : Regex.Replace(CustomerName.Trim(), @"[^a-zA-Z0-9\-_\. ]", "_").Trim();
            if (string.IsNullOrWhiteSpace(safeCustomer))
            {
                safeCustomer = "default-customer";
            }
            var exportDir = Path.Combine(
                Directory.GetCurrentDirectory(),
                "artifacts",
                "customers",
                safeCustomer,
                "exchange");

            var certResult = await _authSession.EnsureExchangeCertificateAsync(
                subjectName: "NewUserAutomation.Exchange",
                exportDirectory: exportDir,
                progress: new Progress<string>(phase => SignInPhaseMessage = phase));
            if (!certResult.Success)
            {
                ReportActionableSetupFailure($"Certificate generation failed. {certResult.ErrorMessage}");
                return;
            }

            ExchangeThumbprint = certResult.Thumbprint;
            ExchangeSetupThumbprint = certResult.Thumbprint;
            var cerPath = certResult.CerPath;

            var setup = await _authSession.EnsureCustomerEnterpriseAppSetupAsync(
                customerName: safeCustomer,
                tenantDomain: TenantDomain,
                certificateCerPath: cerPath,
                certificateThumbprint: ExchangeThumbprint,
                progress: new Progress<string>(phase => SignInPhaseMessage = phase));
            if (!setup.Success)
            {
                ReportActionableSetupFailure(setup.ErrorMessage);
                return;
            }

            ExchangeClientId = setup.ClientId;
            PnPClientId = setup.ClientId;
            ExchangeThumbprint = setup.Thumbprint;
            ExchangeSetupThumbprint = setup.Thumbprint;
            PnPThumbprint = setup.Thumbprint;

            SaveCustomerProfile();
            UpsertCurrentCustomer();

            if (!string.IsNullOrWhiteSpace(setup.ConsentUrl))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = setup.ConsentUrl,
                    UseShellExecute = true
                });
            }

            string sitePermissionDetail = string.Empty;
            if (IsUsableSharePointUrl(SelectedSharePointSiteUrl))
            {
                var siteGrant = await _authSession.EnsureSharePointSiteFullControlForAppAsync(
                    setup.ClientId,
                    SelectedSharePointSiteUrl,
                    TenantDomain,
                    new Progress<string>(phase => SignInPhaseMessage = phase));
                if (!siteGrant.Success)
                {
                    sitePermissionDetail = $" Site grant warning: {Compact(siteGrant.Detail)}";
                }
                else
                {
                    sitePermissionDetail = $" {Compact(siteGrant.Detail)}";
                }
            }

            ExchangeSetupStatus = $"Customer app ready. Client ID: {setup.ClientId}. Complete admin consent in browser, then click Save Customer and Connect All.{sitePermissionDetail}";
            StatusMessage = string.IsNullOrWhiteSpace(setup.ErrorMessage)
                ? $"Customer app setup complete. Next: Save Customer, then Connect All.{sitePermissionDetail}"
                : $"Customer app setup complete with warning: {Compact(setup.ErrorMessage)}";
            SignInPhaseMessage = "[COMPLETED] Customer app setup complete";
        }
        catch (Exception ex)
        {
            ReportActionableSetupFailure(ex.Message);
        }
        finally
        {
            IsAuthBusy = false;
        }
    }

    private void SetupCustomerApp()
    {
        _ = SetupCustomerAppAsync();
    }

    private void RunAuthReadinessCheck()
    {
        _ = RunAuthReadinessCheckAsync();
    }

    private async Task RunAuthReadinessCheckAsync()
    {
        if (IsAuthBusy) return;
        IsAuthBusy = true;
        SignInPhaseMessage = "Running auth readiness checks...";
        try
        {
            var lines = new List<string>();
            var appId = string.IsNullOrWhiteSpace(ExchangeClientId) ? PnPClientId : ExchangeClientId;
            var thumbprint = string.IsNullOrWhiteSpace(ExchangeThumbprint) ? PnPThumbprint : ExchangeThumbprint;

            var tenantReady = !string.IsNullOrWhiteSpace(TenantDomain);
            lines.Add($"{(tenantReady ? "PASS" : "FAIL")} Tenant domain configured");

            var appReady = !string.IsNullOrWhiteSpace(appId);
            lines.Add($"{(appReady ? "PASS" : "FAIL")} App Client ID configured");

            var thumbReady = !string.IsNullOrWhiteSpace(thumbprint);
            lines.Add($"{(thumbReady ? "PASS" : "FAIL")} Certificate thumbprint configured");

            if (thumbReady)
            {
                var certExists = await _authSession.CertificateThumbprintExistsAsync(thumbprint);
                if (!certExists)
                {
                    var detectedExchangeThumb = await _authSession.FindNewestCertificateThumbprintBySubjectAsync("NewUserAutomation.Exchange");
                    if (!string.IsNullOrWhiteSpace(detectedExchangeThumb))
                    {
                        ExchangeThumbprint = detectedExchangeThumb;
                        PnPThumbprint = detectedExchangeThumb;
                        certExists = await _authSession.CertificateThumbprintExistsAsync(detectedExchangeThumb);
                    }
                }
                lines.Add($"{(certExists ? "PASS" : "FAIL")} Certificate exists in CurrentUser\\My");
            }

            var siteReady = IsUsableSharePointUrl(SelectedSharePointSiteUrl);
            lines.Add($"{(siteReady ? "PASS" : "FAIL")} SharePoint site URL valid");

            if (siteReady && appReady)
            {
                var siteGrant = await _authSession.EnsureSharePointSiteFullControlForAppAsync(
                    appId,
                    SelectedSharePointSiteUrl,
                    TenantDomain,
                    new Progress<string>(phase => SignInPhaseMessage = phase));
                lines.Add($"{(siteGrant.Success ? "PASS" : "FAIL")} SharePoint site app permission FullControl");
                lines.Add($"Detail: {Compact(siteGrant.Detail)}");
            }
            else
            {
                lines.Add("INFO Skipped SharePoint site app permission check (missing site URL or app ID).");
            }

            StatusMessage = "Auth readiness check complete.";
            MessageBox.Show(string.Join(Environment.NewLine, lines), "Auth Readiness Check", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        finally
        {
            IsAuthBusy = false;
            SignInPhaseMessage = "Idle";
        }
    }

    private void OpenAdminGrantScript()
    {
        var appId = string.IsNullOrWhiteSpace(ExchangeClientId) ? PnPClientId : ExchangeClientId;
        var siteUrl = SelectedSharePointSiteUrl?.Trim() ?? string.Empty;
        if (!IsUsableSharePointUrl(siteUrl) || string.IsNullOrWhiteSpace(appId))
        {
            StatusMessage = "Auth script requires a valid site URL and app client ID.";
            return;
        }

        var siteUri = new Uri(siteUrl);
        var adminHost = Regex.Replace(siteUri.Host, @"\.sharepoint\.com$", "-admin.sharepoint.com", RegexOptions.IgnoreCase);
        var adminUrl = $"{siteUri.Scheme}://{adminHost}";
        var tenantValue = string.IsNullOrWhiteSpace(TenantDomain) ? siteUri.Host : TenantDomain.Trim();
        var safeSiteUrl = siteUrl.Replace("'", "''", StringComparison.Ordinal);
        var safeAppId = appId.Replace("'", "''", StringComparison.Ordinal);
        var safeAdminUrl = adminUrl.Replace("'", "''", StringComparison.Ordinal);
        var safeTenantValue = tenantValue.Replace("'", "''", StringComparison.Ordinal);
        var scriptPath = Path.Combine(Path.GetTempPath(), $"newuserautomation-sitegrant-{Guid.NewGuid():N}.ps1");
        var script = $$"""
$ErrorActionPreference = 'Stop'
$siteUrl = '{{safeSiteUrl}}'
$appId = '{{safeAppId}}'
$adminUrl = '{{safeAdminUrl}}'
$tenant = '{{safeTenantValue}}'
$authClientId = ''

Import-Module PnP.PowerShell -ErrorAction Stop
try {
    $interactiveApp = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "NewUserAutomation.PnP.Interactive $(Get-Random)" -Tenant $tenant -ErrorAction Stop
    $raw = $interactiveApp | ConvertTo-Json -Depth 20 -Compress
    $idCandidates = @(
        $interactiveApp.AppId, $interactiveApp.ClientId, $interactiveApp.appId, $interactiveApp.clientId,
        $interactiveApp.ApplicationId, $interactiveApp.'Application ID', $interactiveApp.'Application (client) ID',
        $interactiveApp.Id
    ) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }
    $authClientId = ($idCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne 'null' } | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($authClientId) -and $raw -match '[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}') {
        $authClientId = $Matches[0]
    }
    if ([string]::IsNullOrWhiteSpace($authClientId)) {
        throw "Could not resolve interactive auth app client ID from Register-PnPEntraIDAppForInteractiveLogin output."
    }
    Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $authClientId -ErrorAction Stop
} catch {
    throw "Unable to authenticate to SharePoint admin for site permission grant. Details: $($_.Exception.Message)"
}
$existing = @(Get-PnPAzureADAppSitePermission -Site $siteUrl -ErrorAction SilentlyContinue | Where-Object AppId -eq $appId)
$grant = $existing | Select-Object -First 1
if ($grant) {
    Set-PnPAzureADAppSitePermission -Site $siteUrl -PermissionId $grant.Id -Permissions FullControl | Out-Null
    Write-Host "Updated existing site permission to FullControl." -ForegroundColor Green
} else {
    Grant-PnPAzureADAppSitePermission -AppId $appId -DisplayName "NewUserAutomation" -Site $siteUrl -Permissions FullControl | Out-Null
    Write-Host "Granted FullControl site permission." -ForegroundColor Green
}
Write-Host "Done. You can close this window and continue in the app." -ForegroundColor Cyan
""";

        File.WriteAllText(scriptPath, script);
        Process.Start(new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoExit -NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\"",
            UseShellExecute = true
        });
        StatusMessage = "Opened SharePoint admin grant script window.";
    }

    private void CopyExchangeThumbprint()
    {
        if (string.IsNullOrWhiteSpace(ExchangeSetupThumbprint))
        {
            return;
        }

        Clipboard.SetText(ExchangeSetupThumbprint);
        StatusMessage = "Exchange thumbprint copied to clipboard.";
    }

    private async Task ConnectPnPAsync()
    {
        if (IsAuthBusy) return;
        IsAuthBusy = true;
        SignInPhaseMessage = "Connecting SharePoint PnP...";
        try
        {
            var thumbExists = await _authSession.CertificateThumbprintExistsAsync(PnPThumbprint);
            if (!thumbExists)
            {
                var detectedThumb = await _authSession.FindNewestCertificateThumbprintBySubjectAsync("NewUserAutomation.PnP");
                if (!string.IsNullOrWhiteSpace(detectedThumb))
                {
                    PnPThumbprint = detectedThumb;
                    SaveCustomerProfile();
                    UpsertCurrentCustomer();
                    StatusMessage = "Detected and restored SharePoint certificate thumbprint from local cert store.";
                }
            }

            var siteUrl = string.IsNullOrWhiteSpace(SelectedSharePointSiteUrl) ? "https://hresourcing.sharepoint.com" : SelectedSharePointSiteUrl;
            var result = await _authSession.ConnectPnPAsync(siteUrl, PnPClientId, TenantDomain, PnPThumbprint, new Progress<string>(phase => SignInPhaseMessage = phase));
            if (!result.Success)
            {
                var friendly = Compact(result.ErrorMessage);
                PnPStatus = $"Failed: {friendly}";
                IsPnPConnected = false;
                StatusMessage = "SharePoint sign-in failed. Ensure your PnP app Client ID is created and admin-consented.";
                return;
            }

            PnPStatus = $"Connected: {result.TenantId}";
            IsPnPConnected = true;
            StatusMessage = "SharePoint PnP connected.";
            SignInPhaseMessage = "PnP connected";
        }
        finally
        {
            IsAuthBusy = false;
        }
    }

    private async Task SetupPnPAppAsync()
    {
        if (IsAuthBusy) return;
        IsAuthBusy = true;
        SignInPhaseMessage = "Setting up SharePoint app registration...";
        try
        {
            var result = await _authSession.EnsurePnPAppRegistrationAsync(TenantDomain, new Progress<string>(phase => SignInPhaseMessage = phase));
            if (!result.Success)
            {
                SetupStatus = $"PnP app setup failed: {Compact(result.ErrorMessage)}";
                StatusMessage = "Could not set up PnP app registration automatically.";
                return;
            }

            PnPClientId = result.AppId;
            PnPThumbprint = result.Thumbprint;
            if (string.IsNullOrWhiteSpace(TenantDomain))
            {
                TenantDomain = result.TenantDomain;
            }

            SaveCustomerProfile();
            UpsertCurrentCustomer();

            if (!string.IsNullOrWhiteSpace(result.ConsentUrl))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = result.ConsentUrl,
                    UseShellExecute = true
                });
            }

            SetupStatus = "PnP app ready. Admin consent opened. Approve it, then Connect SharePoint.";
            StatusMessage = SetupStatus;

            // Try immediate connect so the user doesn't need extra clicks.
            await ConnectPnPAsync();
        }
        finally
        {
            IsAuthBusy = false;
        }
    }

    private void ParseAndReview()
    {
        var request = _parser.ParseFromFile(SelectedFilePath);
        LoadEditFieldsFromRequest(request);
        _currentRequest = request;
        ApplyEditedValuesInternal("Review ready from parsed form.");
        CurrentStep = WizardStep.Review;
    }

    private void ApplyReviewEdits()
    {
        ApplyEditedValuesInternal("Review refreshed from manual edits.");
    }

    private async void RunPreflight()
    {
        if (IsAuthBusy)
        {
            return;
        }

        IsAuthBusy = true;
        SignInPhaseMessage = "Running preflight checks...";
        try
        {
            if (_currentRequest is null)
            {
                PreflightResult = "No parsed request available.";
                return;
            }

            var editedRequest = BuildRequestFromEdits(_currentRequest.ParseDiagnostics);
            var editedReport = _validation.Validate(editedRequest);
            if (!editedReport.IsValid)
            {
                PreflightResult = $"Preflight blocked: {string.Join(" | ", editedReport.Errors)}";
                return;
            }

            _currentRequest = editedRequest;

            PreflightChecks.Clear();
            AddPreflightSection("Identity");
            var hasFirstLast = !string.IsNullOrWhiteSpace(_currentRequest.FirstName) && !string.IsNullOrWhiteSpace(_currentRequest.LastName);
            AddPreflight("Identity fields present", hasFirstLast, hasFirstLast ? "First and last name present." : "First/last name missing in form.");
            var hasDisplayName = !string.IsNullOrWhiteSpace(_currentRequest.DisplayName);
            AddPreflight("Display name present", hasDisplayName, hasDisplayName ? _currentRequest.DisplayName : "Display name was not parsed.");

        var hasUsername = !string.IsNullOrWhiteSpace(_currentRequest.PreferredUsername);
        AddPreflight("Preferred username (UPN local part) present", hasUsername, hasUsername ? _currentRequest.PreferredUsername : "Preferred username missing.");
        var hasPassword = !string.IsNullOrWhiteSpace(_currentRequest.TemporaryPassword);
        AddPreflight("Temporary password captured", hasPassword, hasPassword ? "Password value parsed from form." : "Temporary password missing from form.");
        var hasJobTitle = !string.IsNullOrWhiteSpace(_currentRequest.JobTitle);
        AddPreflight("Job title captured", hasJobTitle, hasJobTitle ? _currentRequest.JobTitle : "Job title missing from form.");
        var primaryEmailValid = Regex.IsMatch(_currentRequest.PrimaryEmail, @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
        AddPreflight("Primary email format valid", primaryEmailValid, primaryEmailValid ? _currentRequest.PrimaryEmail : "Invalid or missing primary email.");
        if (!string.IsNullOrWhiteSpace(_currentRequest.PreferredUsername) && primaryEmailValid)
        {
            var localPart = _currentRequest.PrimaryEmail.Split('@')[0];
            var preferredUpnLocalPart = _currentRequest.PreferredUsername.Contains('@', StringComparison.Ordinal)
                ? _currentRequest.PreferredUsername.Split('@')[0]
                : _currentRequest.PreferredUsername;
            var matches = string.Equals(localPart, preferredUpnLocalPart, StringComparison.OrdinalIgnoreCase);
            AddPreflight(
                "UPN + email alignment",
                matches,
                matches
                    ? $"Username/UPN local part '{preferredUpnLocalPart}' matches primary email local part '{localPart}'."
                    : $"Username/UPN local part '{preferredUpnLocalPart}' differs from primary email local part '{localPart}'.");
        }
        else
        {
            AddPreflight("UPN + email alignment", false, "Cannot compare due to missing username (UPN local part) or invalid primary email.");
        }
        AddPreflight("Approver provided", !string.IsNullOrWhiteSpace(_currentRequest.RequestApprovedBy), string.IsNullOrWhiteSpace(_currentRequest.RequestApprovedBy) ? "Approver is required." : _currentRequest.RequestApprovedBy);

        AddPreflightSection("Entitlements");
        var duplicateLicenses = _currentRequest.LicenseSkus
            .GroupBy(x => x, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToList();
        AddPreflight("No duplicate licenses", duplicateLicenses.Count == 0, duplicateLicenses.Count == 0 ? "No duplicate license assignments." : $"Duplicates: {string.Join(", ", duplicateLicenses)}");
        foreach (var license in _currentRequest.LicenseSkus.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            AddPreflight($"License item: {license}", true, "Parsed from form.");
        }
        foreach (var group in _currentRequest.GroupAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            AddPreflight($"Group item: {group}", group.Length >= 2, group.Length >= 2 ? "Group target format accepted." : "Group target appears invalid.");
        }

        AddPreflightSection("Connectivity");
        AddPreflight("Graph connected", IsGraphConnected, IsGraphConnected ? GraphStatus : "Connect Graph before execution mode.");
        AddPreflight("Exchange connected", IsExchangeConnected, IsExchangeConnected ? ExchangeStatus : "Connect Exchange to verify mailbox targets.");
        AddPreflight("SharePoint path selected or skipped", SkipSharePointForNow || IsPnPConnected, SkipSharePointForNow ? "Skipped by operator." : PnPStatus);
        AddPreflight("SharePoint site URL format", Uri.TryCreate(SelectedSharePointSiteUrl, UriKind.Absolute, out _), $"URL: {SelectedSharePointSiteUrl}");
        AddPreflight("PnP app details ready", SkipSharePointForNow || (!string.IsNullOrWhiteSpace(PnPClientId) && !string.IsNullOrWhiteSpace(PnPThumbprint) && !string.IsNullOrWhiteSpace(TenantDomain)),
            SkipSharePointForNow ? "Skipped by operator." : "Client ID, thumbprint, and tenant domain are required.");

        AddPreflightSection("Permissions");

        var actions = _dryRunLines
            .Where(line => line.Contains(':'))
            .Select(line => line.Split(':', 2)[0])
            .Select(action => action.Contains('.') ? action.Split('.', 2)[1] : action)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var check = _permissionMatrix.Evaluate(
            selectedActions: actions,
            grantedGraphScopes: ["User.ReadWrite.All", "Group.ReadWrite.All"],
            grantedExchangeRoles: ["Mailbox.Permission.Assign"],
            grantedPnPRights: ["Site.Member.Write"]);

        AddPreflight("Permission matrix satisfied", check.CanExecute, check.CanExecute ? "All required permissions present." : string.Join(" | ", check.MissingPermissions));

        AddPreflightSection("Targets");
        AddPreflight(
            "Mailbox target validation backend",
            IsExchangeConnected,
            IsExchangeConnected
                ? $"Exchange Online session active. {ExchangeStatus}"
                : "Exchange not connected. Mailbox/group target validation cannot run.");
        if (_currentRequest.SharedMailboxAccess.Count > 0 && IsExchangeConnected)
        {
            try
            {
                using var exchangeLookupCts = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var targetLookup = await _authSession.CheckExchangeAccessTargetsAsync(
                    _currentRequest.SharedMailboxAccess,
                    new Progress<string>(phase => SignInPhaseMessage = phase),
                    exchangeLookupCts.Token);
                foreach (var target in _currentRequest.SharedMailboxAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    if (!targetLookup.TryGetValue(target, out var resolution) || !resolution.Exists)
                    {
                        var missingDetail = "Not found in Exchange.";
                        if (resolution is not null && !string.IsNullOrWhiteSpace(resolution.Details))
                        {
                            missingDetail += $" Lookup detail: {Compact(resolution.Details)}";
                        }

                        AddPreflight($"Exchange target exists: {target}", false, missingDetail);
                        continue;
                    }

                    var pass = !string.Equals(resolution.Kind, "OtherRecipient", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(resolution.Kind, "UserMailbox", StringComparison.OrdinalIgnoreCase);
                    var detail = $"Type: {resolution.Kind}. Required action: {resolution.RequiredAction}";
                    AddPreflight($"Exchange target exists: {target}", pass, detail);
                }
            }
            catch (OperationCanceledException)
            {
                foreach (var target in _currentRequest.SharedMailboxAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    AddPreflight($"Exchange target exists: {target}", false, "Exchange target validation timed out after 20 seconds. Preflight continued.");
                }
            }
            catch (Exception ex)
            {
                var detail = $"Exchange validation failed: {Compact(ex.Message)}";
                foreach (var target in _currentRequest.SharedMailboxAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    AddPreflight($"Exchange target exists: {target}", false, detail);
                }
            }
        }
        else if (_currentRequest.SharedMailboxAccess.Count > 0)
        {
            foreach (var target in _currentRequest.SharedMailboxAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                AddPreflight($"Exchange target exists: {target}", false, "Exchange not connected. Connect Exchange and rerun preflight.");
            }
        }
        foreach (var sharePoint in _currentRequest.SharePointAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            var isMappedGroup = sharePoint.StartsWith("SP-Documents-DL-", StringComparison.OrdinalIgnoreCase);
            var isKnownManualSelection = string.Equals(sharePoint, "PAYROLL (ACCOUNTS ONLY)", StringComparison.OrdinalIgnoreCase);
            var valid = sharePoint.Length >= 2 && !sharePoint.Contains('@');
            AddPreflight(
                $"SharePoint access item: {sharePoint}",
                valid && (isMappedGroup || isKnownManualSelection),
                !valid
                    ? "Target item appears invalid."
                    : isMappedGroup
                        ? "Mapped to expected security group format."
                        : isKnownManualSelection
                            ? "Known SharePoint option captured; requires manual mapping to correct security group."
                            : "Does not match expected SP-Documents-DL-* security group naming.");
        }

        if (_currentRequest.RequiresPersonalSharePointFolder)
        {
            AddPreflight(
                "Personal SharePoint folder requested",
                true,
                $"Document library '{_currentRequest.PersonalSharePointFolderName}' and group '{_currentRequest.PersonalSharePointPermissionGroup}' will be created.");
            AddPreflight(
                "Personal library target",
                true,
                $"Create document library '{_currentRequest.PersonalSharePointFolderName}' and apply permissions via '{_currentRequest.PersonalSharePointPermissionGroup}'.");
            AddPreflight(
                "Personal library group naming",
                _currentRequest.PersonalSharePointPermissionGroup.StartsWith("SP-Documents-DL-", StringComparison.OrdinalIgnoreCase)
                && _currentRequest.PersonalSharePointPermissionGroup.EndsWith("-RW", StringComparison.OrdinalIgnoreCase),
                _currentRequest.PersonalSharePointPermissionGroup);
            AddPreflight(
                "Personal library additional user target group",
                !string.IsNullOrWhiteSpace(_currentRequest.PersonalSharePointPermissionGroup),
                $"All additional personal-library users will be added to '{_currentRequest.PersonalSharePointPermissionGroup}'.");
            AddPreflight(
                "SharePoint connection required for personal library creation",
                IsPnPConnected,
                IsPnPConnected
                    ? "PnP connected; personal library and SharePoint group can be created."
                    : "Connect SharePoint (PnP) before execution if personal library creation is required.");
            AddPreflight(
                "Personal library additional members parsed",
                _currentRequest.PersonalSharePointAdditionalMembers.Count > 0,
                _currentRequest.PersonalSharePointAdditionalMembers.Count > 0
                    ? string.Join(", ", _currentRequest.PersonalSharePointAdditionalMembers)
                    : "No additional members found in 'Access over Users SharePoint Folder'.");

            var personalLookupInputs = NormalizeEditList(EditPersonalSharePointAdditionalMembers);
            AddPreflight(
                "Personal library user lookup backend",
                IsGraphConnected,
                IsGraphConnected
                    ? string.IsNullOrWhiteSpace(_graphTenantId)
                        ? $"Microsoft Graph session active. {GraphStatus}"
                        : $"Microsoft Graph session active (TenantId: {_graphTenantId}). {GraphStatus}"
                    : "Graph not connected. User identity validation cannot run.");
            if (personalLookupInputs.Count > 0 && IsGraphConnected)
            {
                var priorSelections = PersonalFolderUserMatchChoices
                    .Where(choice => !string.IsNullOrWhiteSpace(choice.SelectedPrincipal))
                    .ToDictionary(choice => choice.SourceIdentity, choice => choice.SelectedPrincipal, StringComparer.OrdinalIgnoreCase);
                var userLookup = await _authSession.FindDirectoryUserMatchesAsync(
                    personalLookupInputs,
                    _graphTenantId,
                    new Progress<string>(phase => SignInPhaseMessage = phase));
                var foundCount = 0;
                var missingCount = 0;
                var ambiguousCount = 0;
                PersonalFolderUserMatchChoices.Clear();
                foreach (var identity in personalLookupInputs.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    var matches = userLookup.TryGetValue(identity, out var value)
                        ? value
                        : [];

                    if (matches.Count > 0)
                    {
                        var choice = new DirectoryUserMatchChoice(identity, matches);
                        if (priorSelections.TryGetValue(identity, out var selected) && choice.MatchOptions.Any(option => option.Equals(selected, StringComparison.OrdinalIgnoreCase)))
                        {
                            choice.SelectedPrincipal = selected;
                        }
                        else if (TryGetExactIdentityMatchOption(identity, matches, out var exactOption))
                        {
                            choice.SelectedPrincipal = exactOption;
                        }
                        else if (matches.Count == 1 && choice.MatchOptions.Count == 1)
                        {
                            choice.SelectedPrincipal = choice.MatchOptions[0];
                        }
                        PersonalFolderUserMatchChoices.Add(choice);

                        if (matches.Count > 1 && string.IsNullOrWhiteSpace(choice.SelectedPrincipal))
                        {
                            ambiguousCount++;
                            AddPreflight(
                                $"Personal library member match selection required: {identity}",
                                false,
                                $"Found {matches.Count} first-name candidate account(s). Choose the correct account in Review and rerun preflight.");
                        }
                        else
                        {
                            foundCount++;
                            var selectedLabel = !string.IsNullOrWhiteSpace(choice.SelectedPrincipal)
                                ? choice.SelectedPrincipal
                                : FormatDirectoryMatch(matches[0]);
                            AddPreflight(
                                $"Personal library member resolved: {identity}",
                                true,
                                $"Using account '{selectedLabel}'.");
                        }
                    }
                    else
                    {
                        missingCount++;
                        AddPreflight(
                            $"Personal library member exists: {identity}",
                            false,
                            "User identity not found. Update the additional users list before sign-off.");
                    }
                }

                OnPropertyChanged(nameof(PersonalFolderUserMatchChoices));

                AddPreflight(
                    "Personal library directory validation executed",
                    ambiguousCount == 0,
                    $"Checked {personalLookupInputs.Count} user(s); {foundCount} resolved, {ambiguousCount} awaiting choice, {missingCount} missing.");
            }
            else if (personalLookupInputs.Count > 0)
            {
                PersonalFolderUserMatchChoices.Clear();
                OnPropertyChanged(nameof(PersonalFolderUserMatchChoices));
                foreach (var identity in personalLookupInputs.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    AddPreflight(
                        $"Personal library member exists: {identity}",
                        false,
                        "Graph not connected. Connect Graph and rerun preflight.");
                }
                AddPreflight(
                    "Personal library directory validation executed",
                    false,
                    "Skipped because Graph is not connected.");
            }
            else
            {
                PersonalFolderUserMatchChoices.Clear();
                OnPropertyChanged(nameof(PersonalFolderUserMatchChoices));
                AddPreflight(
                    "Personal library directory validation executed",
                    false,
                    "No additional users were parsed to validate.");
            }
        }

        AddPreflightSection("Parser Diagnostics");
        if (_currentRequest.ParseDiagnostics.Count == 0)
        {
            AddPreflight("Diagnostics available", true, "No parser diagnostics emitted.");
        }
        else
        {
            if (!ShowParserDiagnosticsDetails)
            {
                AddPreflight(
                    "Parser diagnostics collapsed",
                    true,
                    $"{_currentRequest.ParseDiagnostics.Count} diagnostic item(s) hidden. Click '{ParserDiagnosticsToggleText}' to expand.");
            }
            else
            {
                foreach (var message in _currentRequest.ParseDiagnostics.OrderBy(x => x, StringComparer.Ordinal))
                {
                    var isFailure = message.Contains("missing", StringComparison.OrdinalIgnoreCase) || message.Contains("no match", StringComparison.OrdinalIgnoreCase);
                    AddPreflight($"Parser: {message}", !isFailure, isFailure ? "Check source form field labels/content." : "Diagnostic info.");
                }
            }
        }

            var allPassed = PreflightChecks.All(c => c.IsPass);
            PreflightResult = allPassed
                ? "Preflight passed. Dry-run is ready for operator sign-off."
                : "Preflight blocked. Resolve failed checks, then rerun preflight.";
            RefreshCommandStates();
        }
        catch (Exception ex)
        {
            PreflightChecks.Add(new PreflightCheckItem("FAIL", "Preflight runtime error", Compact(ex.Message), false));
            PreflightResult = "Preflight blocked due to runtime error. See failed check for details.";
            StatusMessage = $"Preflight error: {Compact(ex.Message)}";
            RefreshCommandStates();
        }
        finally
        {
            IsAuthBusy = false;
            if (string.IsNullOrWhiteSpace(SignInPhaseMessage) || SignInPhaseMessage.StartsWith("Running preflight", StringComparison.OrdinalIgnoreCase))
            {
                SignInPhaseMessage = "Preflight complete";
            }
        }
    }

    private void StartSafeLiveRun()
    {
        if (_currentRequest is null)
        {
            LiveRunStatus = "[FAILED] No parsed request available.";
            return;
        }

        var editedRequest = BuildRequestFromEdits(_currentRequest.ParseDiagnostics);
        var report = _validation.Validate(editedRequest);
        if (!report.IsValid)
        {
            LiveRunStatus = $"[FAILED] Live run blocked: {string.Join(" | ", report.Errors)}";
            return;
        }
        if (PreflightChecks.Count == 0 || PreflightChecks.Any(check => !check.IsPass))
        {
            LiveRunStatus = "[FAILED] Run preflight first and resolve all failures before safe live execution.";
            return;
        }

        _currentRequest = editedRequest;
        var checkpointPath = GetLiveRunCheckpointPath(_currentRequest.Upn);
        if (File.Exists(checkpointPath))
        {
            try
            {
                var existingJson = File.ReadAllText(checkpointPath);
                var existingState = JsonSerializer.Deserialize<LiveRunState>(existingJson);
                if (existingState is not null && existingState.Steps.Count > 0)
                {
                    var resume = MessageBox.Show(
                        $"Existing live run checkpoint found for {_currentRequest.Upn}. Resume from last completed step?",
                        "Resume Live Run",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    _liveRunState = resume == MessageBoxResult.Yes ? existingState : BuildNewLiveRunState(_currentRequest);
                }
            }
            catch
            {
                _liveRunState = BuildNewLiveRunState(_currentRequest);
            }
        }

        _liveRunState ??= BuildNewLiveRunState(_currentRequest);
        SyncLiveStepsFromState();
        SaveLiveRunState();
        LiveRunLogPath = GetLiveRunLogPath(_liveRunState.UserUpn);
        LogLiveRunEvent("LiveRunStarted", null, "Safe live run initialized or resumed.");
        LiveRunStatus = "[READY] Safe live run initialized. Use 'Run Next Live Step' or 'Run All Remaining Steps'.";
        UpdateLiveRunProgress();
    }

    private async void ExecuteNextLiveStep()
    {
        if (_currentRequest is null)
        {
            LiveRunStatus = "[FAILED] No parsed request available.";
            return;
        }

        if (_liveRunState is null)
        {
            StartSafeLiveRun();
            if (_liveRunState is null)
            {
                return;
            }
        }

        var nextStep = _liveRunState.Steps.FirstOrDefault(step => !step.IsCompleted);
        if (nextStep is null)
        {
            LiveRunStatus = "[COMPLETED] Live run complete. All steps finished.";
            return;
        }

        var confirm = MessageBox.Show(
            string.Equals(nextStep.Status, "Failed", StringComparison.OrdinalIgnoreCase)
                ? $"Retry failed step?\n\n{nextStep.Description}\n\nLast error: {nextStep.Detail}"
                : $"Execute next step?\n\n{nextStep.Description}",
            "Confirm Live Step",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            LiveRunStatus = "[PAUSED] Live run paused by operator.";
            return;
        }

        IsAuthBusy = true;
        SignInPhaseMessage = $"[RUNNING] Live step: {nextStep.Description}";
        try
        {
            nextStep.Status = "InProgress";
            nextStep.UpdatedUtc = DateTimeOffset.UtcNow;
            SyncLiveStepsFromState();
            LogLiveRunEvent("StepStarted", nextStep, "Operator confirmed execution.");

            var result = await ExecuteLiveStepInternal(nextStep, _currentRequest);
            nextStep.Status = result.Success ? "Completed" : "Failed";
            nextStep.Detail = result.Detail;
            nextStep.UpdatedUtc = DateTimeOffset.UtcNow;
            _liveRunState.UpdatedUtc = DateTimeOffset.UtcNow;
            SaveLiveRunState();
            SyncLiveStepsFromState();
            LogLiveRunEvent(result.Success ? "StepCompleted" : "StepFailed", nextStep, result.Detail);

            if (result.Success)
            {
                LiveRunStatus = $"[COMPLETED] Step completed: {nextStep.Description}";
                var remaining = _liveRunState.Steps.Count(step => !step.IsCompleted);
                if (remaining == 0)
                {
                    var skipped = _liveRunState.Steps.Count(step => string.Equals(step.Status, "Skipped", StringComparison.OrdinalIgnoreCase));
                    LiveRunStatus = skipped > 0
                        ? $"[COMPLETED] Live run complete with {skipped} skipped step(s). Review skipped items before sign-off."
                        : "[COMPLETED] Live run complete. All steps finished.";
                }
            }
            else
            {
                var actionable = BuildLiveRunFailureDetail(result.Detail);
                LiveRunStatus = $"[FAILED] Step failed: {nextStep.Description}. {actionable}";
                MessageBox.Show(
                    $"Step failed and the live run has stopped.\n\n{nextStep.Description}\n\n{actionable}\n\nFix the issue, then click 'Run Next Live Step' to retry this failed step.",
                    "Live Run Stopped",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            UpdateLiveRunProgress();
        }
        finally
        {
            IsAuthBusy = false;
            RefreshCommandStates();
        }
    }

    private void SkipLiveStep()
    {
        if (_liveRunState is null)
        {
            return;
        }

        var nextStep = _liveRunState.Steps.FirstOrDefault(step => !step.IsCompleted);
        if (nextStep is null)
        {
            LiveRunStatus = "[PAUSED] No remaining steps to skip.";
            return;
        }

        var confirm = MessageBox.Show(
            $"Skip this step?\n\n{nextStep.Description}\n\nThis marks the step as Skipped and it will not be executed unless you reset the live run.",
            "Skip Live Step",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        nextStep.Status = "Skipped";
        nextStep.Detail = "Skipped by operator.";
        nextStep.UpdatedUtc = DateTimeOffset.UtcNow;
        _liveRunState.UpdatedUtc = DateTimeOffset.UtcNow;
        SaveLiveRunState();
        SyncLiveStepsFromState();
        LogLiveRunEvent("StepSkipped", nextStep, nextStep.Detail);

        var remaining = _liveRunState.Steps.Count(step => !step.IsCompleted);
        if (remaining == 0)
        {
            var skipped = _liveRunState.Steps.Count(step => string.Equals(step.Status, "Skipped", StringComparison.OrdinalIgnoreCase));
            LiveRunStatus = skipped > 0
                ? $"[COMPLETED] Live run complete with {skipped} skipped step(s). Review skipped items before sign-off."
                : "[COMPLETED] Live run complete. All steps finished.";
        }
        else
        {
            LiveRunStatus = $"[PAUSED] Step skipped: {nextStep.Description}";
        }
        UpdateLiveRunProgress();
    }

    private void BackToStart()
    {
        CurrentStep = WizardStep.SignIn;
        SelectedFilePath = string.Empty;
        _currentRequest = null;
        _dryRunLines = [];
        OnPropertyChanged(nameof(DryRunLines));
        PreflightChecks.Clear();
        PreflightResult = "Preflight not run.";
        PersonalFolderUserMatchChoices.Clear();
        OnPropertyChanged(nameof(PersonalFolderUserMatchChoices));
        _liveRunState = null;
        LiveRunSteps.Clear();
        LiveRunSummarySteps.Clear();
        LiveRunLogPath = string.Empty;
        LiveRunStatus = "[READY] Ready for a new user onboarding run.";
        UpdateLiveRunProgress();
        SignInPhaseMessage = "Idle";
    }

    private static void CloseApplication()
    {
        Application.Current.Shutdown();
    }

    private async void ExecuteAllLiveSteps()
    {
        if (_currentRequest is null)
        {
            LiveRunStatus = "[FAILED] No parsed request available.";
            return;
        }

        if (_liveRunState is null)
        {
            StartSafeLiveRun();
            if (_liveRunState is null)
            {
                return;
            }
        }

        var total = _liveRunState.Steps.Count;
        var remainingAtStart = _liveRunState.Steps.Count(step => !step.IsCompleted);
        if (remainingAtStart == 0)
        {
            LiveRunStatus = "[COMPLETED] Live run complete. All steps finished.";
            UpdateLiveRunProgress();
            return;
        }

        var confirm = MessageBox.Show(
            $"Run all remaining steps in order?\n\nRemaining: {remainingAtStart} of {total}\n\nFailed steps will be auto-marked as Skipped and execution will continue.",
            "Run All Remaining Steps",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        IsAuthBusy = true;
        var autoSkippedFailures = 0;
        try
        {
            while (true)
            {
                var nextStep = _liveRunState.Steps.FirstOrDefault(step => !step.IsCompleted);
                if (nextStep is null)
                {
                    break;
                }

                var done = _liveRunState.Steps.Count(step => step.IsCompleted);
                SignInPhaseMessage = $"Running step {done + 1}/{total}: {nextStep.Description}";
                LiveRunStatus = $"[RUNNING] Run-all ({done + 1}/{total}): {nextStep.Description}";

                nextStep.Status = "InProgress";
                nextStep.UpdatedUtc = DateTimeOffset.UtcNow;
                SyncLiveStepsFromState();
                LogLiveRunEvent("StepStarted", nextStep, "Run-all execution.");

                var result = await ExecuteLiveStepInternal(nextStep, _currentRequest);
                nextStep.Status = result.Success ? "Completed" : "Failed";
                nextStep.Detail = result.Detail;
                nextStep.UpdatedUtc = DateTimeOffset.UtcNow;
                _liveRunState.UpdatedUtc = DateTimeOffset.UtcNow;
                SaveLiveRunState();
                SyncLiveStepsFromState();
                LogLiveRunEvent(result.Success ? "StepCompleted" : "StepFailed", nextStep, result.Detail);

                if (!result.Success)
                {
                    nextStep.Status = "Skipped";
                    nextStep.Detail = $"Auto-skipped during run-all after failure: {result.Detail}";
                    nextStep.UpdatedUtc = DateTimeOffset.UtcNow;
                    _liveRunState.UpdatedUtc = DateTimeOffset.UtcNow;
                    SaveLiveRunState();
                    SyncLiveStepsFromState();
                    LogLiveRunEvent("StepSkipped", nextStep, nextStep.Detail);
                    autoSkippedFailures++;
                }

                UpdateLiveRunProgress();
            }

            var skipped = _liveRunState.Steps.Count(step => string.Equals(step.Status, "Skipped", StringComparison.OrdinalIgnoreCase));
            LiveRunStatus = skipped > 0
                ? $"[COMPLETED] Run-all complete. {skipped} step(s) skipped ({autoSkippedFailures} due to failures)."
                : "[COMPLETED] Run-all complete. All steps finished.";
        }
        finally
        {
            IsAuthBusy = false;
            SignInPhaseMessage = "Live run complete";
            RefreshCommandStates();
            UpdateLiveRunProgress();
        }
    }

    private void FinishLiveRun()
    {
        if (_liveRunState is null || _liveRunState.Steps.Count == 0)
        {
            LiveRunStatus = "[FAILED] No live run data available to summarize.";
            return;
        }

        if (_liveRunState.Steps.Any(step => !step.IsCompleted))
        {
            LiveRunStatus = "[PAUSED] All steps must be completed or skipped before finishing.";
            return;
        }

        LiveRunSummarySteps.Clear();
        foreach (var step in _liveRunState.Steps
                     .OrderByDescending(step => string.Equals(step.Status, "Skipped", StringComparison.OrdinalIgnoreCase))
                     .ThenBy(step => step.Description, StringComparer.OrdinalIgnoreCase))
        {
            LiveRunSummarySteps.Add(step);
        }

        OnPropertyChanged(nameof(LiveRunSummarySteps));
        OnPropertyChanged(nameof(LiveRunSummaryHeadline));
        CurrentStep = WizardStep.Summary;
    }

    private void ResetLiveRun()
    {
        if (_currentRequest is null)
        {
            return;
        }

        var confirm = MessageBox.Show(
            "Reset checkpoint and rebuild safe live run steps from current review values?",
            "Reset Live Run",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        var path = GetLiveRunCheckpointPath(_currentRequest.Upn);
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }

        _liveRunState = BuildNewLiveRunState(_currentRequest);
        SyncLiveStepsFromState();
        SaveLiveRunState();
        LiveRunLogPath = GetLiveRunLogPath(_liveRunState.UserUpn);
        LogLiveRunEvent("LiveRunReset", null, "Operator reset safe live run checkpoint.");
        LiveRunStatus = "[READY] Live run reset. Ready to execute from step 1.";
        UpdateLiveRunProgress();
    }

    private async Task<LiveExecutionResult> ExecuteLiveStepInternal(LiveRunStepState step, NewUserRequest request)
    {
        return step.Command switch
        {
            "CreateUser" => await _authSession.EnsureUserExistsAsync(request),
            "EnsurePersonalGroup" => await _authSession.EnsureSecurityGroupAsync(step.Argument),
            "CreatePersonalFolderAtRoot" => await ExecuteCreatePersonalFolderStep(step.Argument),
            "ApplyPersonalFolderPermissions" => await ExecuteApplyPersonalFolderPermissionsStep(step.Argument),
            "AssignLicense" => await _authSession.AssignLicenseBySkuPartNumberAsync(request.Upn, step.Argument),
            "AddGroupMembership" => await _authSession.AddUserToGroupByNameAsync(request.Upn, step.Argument),
            "AddNamedMemberToGroup" => await ExecuteNamedMemberGroupStep(step.Argument),
            "GrantExchangeAccessBatch" => await ExecuteExchangeBatchStep(request.Upn, step.Argument),
            "GrantExchangeAccess" => await _authSession.GrantExchangeTargetAccessAsync(request.Upn, step.Argument),
            _ => new LiveExecutionResult(false, $"Unknown live step command: {step.Command}")
        };
    }

    private async Task<LiveExecutionResult> ExecuteNamedMemberGroupStep(string argument)
    {
        var parts = argument.Split('|', 2, StringSplitOptions.TrimEntries);
        if (parts.Length != 2 || string.IsNullOrWhiteSpace(parts[0]) || string.IsNullOrWhiteSpace(parts[1]))
        {
            return new LiveExecutionResult(false, $"Invalid named-member argument: {argument}");
        }

        return await _authSession.AddUserToGroupByNameAsync(parts[1], parts[0]);
    }

    private async Task<LiveExecutionResult> ExecuteCreatePersonalFolderStep(string argument)
    {
        if (string.IsNullOrWhiteSpace(argument))
        {
            return new LiveExecutionResult(false, "Invalid personal library step argument.");
        }

        if (string.IsNullOrWhiteSpace(SelectedSharePointSiteUrl) ||
            string.IsNullOrWhiteSpace(PnPClientId) ||
            string.IsNullOrWhiteSpace(TenantDomain) ||
            string.IsNullOrWhiteSpace(PnPThumbprint))
        {
            return new LiveExecutionResult(false, "Missing SharePoint connection details (site URL/client ID/tenant/thumbprint).");
        }

        return await _authSession.EnsurePersonalFolderAtSiteRootAsync(
            SelectedSharePointSiteUrl,
            argument,
            PnPClientId,
            TenantDomain,
            PnPThumbprint);
    }

    private async Task<LiveExecutionResult> ExecuteExchangeBatchStep(string userUpn, string argument)
    {
        var targets = argument
            .Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return await _authSession.GrantExchangeTargetsAccessAsync(userUpn, targets);
    }

    private async Task<LiveExecutionResult> ExecuteApplyPersonalFolderPermissionsStep(string argument)
    {
        var parts = argument.Split('|', 2, StringSplitOptions.TrimEntries);
        if (parts.Length != 2 || string.IsNullOrWhiteSpace(parts[0]) || string.IsNullOrWhiteSpace(parts[1]))
        {
            return new LiveExecutionResult(false, $"Invalid personal-folder permission argument: {argument}");
        }

        if (string.IsNullOrWhiteSpace(SelectedSharePointSiteUrl) ||
            string.IsNullOrWhiteSpace(PnPClientId) ||
            string.IsNullOrWhiteSpace(TenantDomain) ||
            string.IsNullOrWhiteSpace(PnPThumbprint))
        {
            return new LiveExecutionResult(false, "Missing SharePoint connection details (site URL/client ID/tenant/thumbprint).");
        }

        return await _authSession.ApplyPersonalFolderPermissionsAsync(
            SelectedSharePointSiteUrl,
            parts[0],
            parts[1],
            PnPClientId,
            TenantDomain,
            PnPThumbprint);
    }

    private LiveRunState BuildNewLiveRunState(NewUserRequest request)
    {
        var state = new LiveRunState
        {
            UserUpn = request.Upn,
            OperatorAccount = _authSession.GraphAccount,
            TenantId = _authSession.GraphTenantId,
            StartedUtc = DateTimeOffset.UtcNow,
            UpdatedUtc = DateTimeOffset.UtcNow
        };

        var steps = new List<LiveRunStepState>
        {
            NewLiveStep("create-user", "Create user", "CreateUser", request.Upn)
        };

        if (request.RequiresPersonalSharePointFolder && !string.IsNullOrWhiteSpace(request.PersonalSharePointPermissionGroup))
        {
            steps.Add(NewLiveStep(
                "ensure-personal-group",
                $"Create personal library security group '{request.PersonalSharePointPermissionGroup}'",
                "EnsurePersonalGroup",
                request.PersonalSharePointPermissionGroup));
            steps.Add(NewLiveStep(
                "create-personal-folder-root",
                $"Create personal document library '{request.PersonalSharePointFolderName}'",
                "CreatePersonalFolderAtRoot",
                request.PersonalSharePointFolderName));
            steps.Add(NewLiveStep(
                "apply-personal-folder-permissions",
                $"Apply library permissions for '{request.PersonalSharePointPermissionGroup}' on '{request.PersonalSharePointFolderName}'",
                "ApplyPersonalFolderPermissions",
                request.PersonalSharePointFolderName + "|" + request.PersonalSharePointPermissionGroup));
        }

        steps.AddRange(request.LicenseSkus
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .Select(sku => NewLiveStep($"license-{sku}", $"Assign license '{sku}'", "AssignLicense", sku)));

        steps.AddRange(request.GroupAccess
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .Select(group => NewLiveStep($"group-{group}", $"Add user to group '{group}'", "AddGroupMembership", group)));

        steps.AddRange(request.SharePointAccess
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .Select(group => NewLiveStep($"spgroup-{group}", $"Add user to SharePoint access group '{group}'", "AddGroupMembership", group)));

        if (!string.IsNullOrWhiteSpace(request.PersonalSharePointPermissionGroup))
        {
            steps.Add(NewLiveStep(
                $"personal-owner-{request.PersonalSharePointPermissionGroup}",
                $"Add new user to personal library group '{request.PersonalSharePointPermissionGroup}'",
                "AddGroupMembership",
                request.PersonalSharePointPermissionGroup));

            steps.AddRange(request.PersonalSharePointAdditionalMembers
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .Select(member => NewLiveStep(
                    $"personal-member-{member}",
                    $"Add personal-library member '{member}' to '{request.PersonalSharePointPermissionGroup}'",
                    "AddNamedMemberToGroup",
                    request.PersonalSharePointPermissionGroup + "|" + member)));
        }

        if (request.SharedMailboxAccess.Count > 0)
        {
            steps.AddRange(request.SharedMailboxAccess
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(target => NewLiveStep(
                    $"exchange-{target}",
                    $"Grant Exchange access for '{target}'",
                    "GrantExchangeAccess",
                    target)));
        }

        // Deduplicate by stable Id in case of duplicate form entries.
        state.Steps = steps
            .GroupBy(step => step.Id, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();

        return state;
    }

    private static LiveRunStepState NewLiveStep(string id, string description, string command, string argument)
        => new()
        {
            Id = id,
            Description = description,
            Command = command,
            Argument = argument,
            Status = "Pending",
            UpdatedUtc = DateTimeOffset.UtcNow
        };

    private string GetLiveRunCheckpointPath(string upn)
    {
        var safe = Regex.Replace(upn, @"[^a-zA-Z0-9\._-]", "_");
        var folder = Path.Combine(AppContext.BaseDirectory, "artifacts", "live-runs");
        Directory.CreateDirectory(folder);
        return Path.Combine(folder, $"{safe}.json");
    }

    private string GetLiveRunLogPath(string upn)
    {
        var safe = Regex.Replace(upn, @"[^a-zA-Z0-9\._-]", "_");
        var folder = Path.Combine(AppContext.BaseDirectory, "artifacts", "live-runs");
        Directory.CreateDirectory(folder);
        return Path.Combine(folder, $"{safe}.log.jsonl");
    }

    private void SaveLiveRunState()
    {
        if (_liveRunState is null)
        {
            return;
        }

        try
        {
            _liveRunState.UpdatedUtc = DateTimeOffset.UtcNow;
            var json = JsonSerializer.Serialize(_liveRunState, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(GetLiveRunCheckpointPath(_liveRunState.UserUpn), json);
        }
        catch
        {
        }
    }

    private void SyncLiveStepsFromState()
    {
        LiveRunSteps.Clear();
        if (_liveRunState is null)
        {
            return;
        }

        LiveRunLogPath = GetLiveRunLogPath(_liveRunState.UserUpn);

        foreach (var step in _liveRunState.Steps)
        {
            LiveRunSteps.Add(step);
        }

        OnPropertyChanged(nameof(LiveRunSteps));
        UpdateLiveRunProgress();
        RefreshCommandStates();
    }

    private void LogLiveRunEvent(string eventType, LiveRunStepState? step, string detail)
    {
        if (_liveRunState is null)
        {
            return;
        }

        var checkpointPath = GetLiveRunCheckpointPath(_liveRunState.UserUpn);
        var logPath = GetLiveRunLogPath(_liveRunState.UserUpn);
        try
        {
            _liveRunLogger.Append(logPath, new LiveRunLogEvent
            {
                RunId = _liveRunState.RunId,
                EventType = eventType,
                UserUpn = _liveRunState.UserUpn,
                TenantId = _liveRunState.TenantId,
                OperatorAccount = _liveRunState.OperatorAccount,
                StepId = step?.Id ?? string.Empty,
                StepDescription = step?.Description ?? string.Empty,
                Command = step?.Command ?? string.Empty,
                Argument = step?.Argument ?? string.Empty,
                Status = step?.Status ?? string.Empty,
                Detail = detail,
                CheckpointPath = checkpointPath
            });
        }
        catch
        {
            // Logging must not block provisioning flow.
        }
    }

    private void RefreshCommandStates()
    {
        ContinueToUploadCommand.RaiseCanExecuteChanged();
        ContinueToReviewCommand.RaiseCanExecuteChanged();
        ContinueToLiveRunCommand.RaiseCanExecuteChanged();
        BackToReviewCommand.RaiseCanExecuteChanged();
        FinishLiveRunCommand.RaiseCanExecuteChanged();
        BackToLiveRunCommand.RaiseCanExecuteChanged();
        BackToStartCommand.RaiseCanExecuteChanged();
        CloseApplicationCommand.RaiseCanExecuteChanged();
        RunPreflightCommand.RaiseCanExecuteChanged();
        ConnectGraphCommand.RaiseCanExecuteChanged();
        QuickConnectCommand.RaiseCanExecuteChanged();
        ConnectExchangeCommand.RaiseCanExecuteChanged();
        SetupExchangeCertCommand.RaiseCanExecuteChanged();
        SetupCustomerAppCommand.RaiseCanExecuteChanged();
        AuthReadinessCheckCommand.RaiseCanExecuteChanged();
        OpenAdminGrantScriptCommand.RaiseCanExecuteChanged();
        CopyExchangeThumbprintCommand.RaiseCanExecuteChanged();
        ConnectPnPCommand.RaiseCanExecuteChanged();
        SetupPnPAppCommand.RaiseCanExecuteChanged();
        SaveCustomerCommand.RaiseCanExecuteChanged();
        DeleteCustomerCommand.RaiseCanExecuteChanged();
        EditCustomerCommand.RaiseCanExecuteChanged();
        ApplyReviewEditsCommand.RaiseCanExecuteChanged();
        StartSafeLiveRunCommand.RaiseCanExecuteChanged();
        ExecuteNextLiveStepCommand.RaiseCanExecuteChanged();
        ExecuteAllLiveStepsCommand.RaiseCanExecuteChanged();
        SkipLiveStepCommand.RaiseCanExecuteChanged();
        ResetLiveRunCommand.RaiseCanExecuteChanged();
    }

    private void UpdateLiveRunProgress()
    {
        if (_liveRunState is null || _liveRunState.Steps.Count == 0)
        {
            LiveRunProgressPercent = 0;
            LiveRunProgressText = "Progress: 0/0";
            return;
        }

        var total = _liveRunState.Steps.Count;
        var done = _liveRunState.Steps.Count(step => step.IsCompleted);
        LiveRunProgressPercent = (int)Math.Round((double)done * 100 / total);
        LiveRunProgressText = $"Progress: {done}/{total} ({LiveRunProgressPercent}%)";
    }

    private void LoadEditFieldsFromRequest(NewUserRequest request)
    {
        EditFirstName = request.FirstName;
        EditLastName = request.LastName;
        EditDisplayName = request.DisplayName;
        EditPreferredUsername = request.PreferredUsername;
        EditTemporaryPassword = request.TemporaryPassword;
        EditJobTitle = request.JobTitle;
        EditPrimaryEmail = request.PrimaryEmail;
        EditSecondaryEmail = request.SecondaryEmail;
        EditLicenseSkus = string.Join("; ", request.LicenseSkus);
        EditGroupAccess = string.Join("; ", request.GroupAccess);
        EditSharedMailboxAccess = string.Join("; ", request.SharedMailboxAccess);
        EditSharePointAccess = string.Join("; ", request.SharePointAccess);
        EditPersonalSharePointAdditionalMembers = string.Join("; ", request.PersonalSharePointAdditionalMembers);
        EditSpecialRequirements = request.SpecialRequirements;
        EditRequestApprovedBy = request.RequestApprovedBy;
    }

    private NewUserRequest BuildRequestFromEdits(IReadOnlyList<string> diagnostics)
    {
        var existing = _currentRequest;
        var resolvedPersonalMembers = ResolvePersonalMemberSelections(NormalizeEditList(EditPersonalSharePointAdditionalMembers));
        return new NewUserRequest
        {
            FirstName = EditFirstName.Trim(),
            LastName = EditLastName.Trim(),
            DisplayName = EditDisplayName.Trim(),
            PreferredUsername = EditPreferredUsername.Trim().ToLowerInvariant(),
            TemporaryPassword = EditTemporaryPassword.Trim(),
            JobTitle = EditJobTitle.Trim(),
            PrimaryEmail = EditPrimaryEmail.Trim().ToLowerInvariant(),
            SecondaryEmail = EditSecondaryEmail.Trim().ToLowerInvariant(),
            LicenseSkus = NormalizeEditList(EditLicenseSkus, forceUpperUnderscore: true),
            GroupAccess = NormalizeEditList(EditGroupAccess),
            SharedMailboxAccess = NormalizeEditList(EditSharedMailboxAccess),
            SharePointAccess = NormalizeEditList(EditSharePointAccess),
            RequiresPersonalSharePointFolder = existing?.RequiresPersonalSharePointFolder ?? false,
            PersonalSharePointFolderName = existing?.PersonalSharePointFolderName ?? string.Empty,
            PersonalSharePointPermissionGroup = existing?.PersonalSharePointPermissionGroup ?? string.Empty,
            PersonalSharePointAdditionalMembers = resolvedPersonalMembers,
            SpecialRequirements = EditSpecialRequirements.Trim(),
            RequestApprovedBy = EditRequestApprovedBy.Trim(),
            ParseDiagnostics = diagnostics
        };
    }

    private List<string> ResolvePersonalMemberSelections(IReadOnlyList<string> members)
    {
        var selectedMap = PersonalFolderUserMatchChoices
            .Where(choice => !string.IsNullOrWhiteSpace(choice.SelectedPrincipal))
            .ToDictionary(choice => choice.SourceIdentity, choice => choice.SelectedPrincipal!, StringComparer.OrdinalIgnoreCase);

        return members
            .Select(member =>
            {
                if (!selectedMap.TryGetValue(member, out var selected))
                {
                    return member;
                }

                var extracted = ExtractUpnFromMatchLabel(selected);
                return string.IsNullOrWhiteSpace(extracted) ? member : extracted;
            })
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string ExtractUpnFromMatchLabel(string selected)
    {
        if (string.IsNullOrWhiteSpace(selected))
        {
            return string.Empty;
        }

        var start = selected.LastIndexOf('<');
        var end = selected.LastIndexOf('>');
        if (start >= 0 && end > start)
        {
            return selected[(start + 1)..end].Trim();
        }

        return selected.Trim();
    }

    private void ApplyEditedValuesInternal(string successMessagePrefix)
    {
        var existingDiagnostics = _currentRequest?.ParseDiagnostics ?? [];
        var request = BuildRequestFromEdits(existingDiagnostics);
        _currentRequest = request;

        var report = _validation.Validate(request);
        if (!report.IsValid)
        {
            StatusMessage = $"Review blocked: {string.Join(" | ", report.Errors)}";
            _dryRunLines = [];
            OnPropertyChanged(nameof(DryRunLines));
            return;
        }

        var steps = _dryRunPlanner.BuildPlan(request);
        _dryRunLines = [];
        foreach (var group in steps.GroupBy(step => step.Action.Split('.', 2)[0], StringComparer.OrdinalIgnoreCase))
        {
            _dryRunLines.Add($"--- {group.Key} ---");
            foreach (var step in group)
            {
                _dryRunLines.Add($"{step.Action}: {step.Description}");
            }
        }
        OnPropertyChanged(nameof(DryRunLines));

        var parseSummary = request.ParseDiagnostics.Count == 0
            ? "Parser completed."
            : $"Parser completed with {request.ParseDiagnostics.Count} diagnostic message(s).";

        StatusMessage = report.Warnings.Count > 0
            ? $"{successMessagePrefix} Warnings: {string.Join(" | ", report.Warnings)} {parseSummary}"
            : $"{successMessagePrefix} Review ready for preflight. {parseSummary}";

        PreflightResult = "Preflight not run.";
        PreflightChecks.Clear();
    }

    private static List<string> NormalizeEditList(string input, bool forceUpperUnderscore = false)
    {
        var parts = input
            .Replace("\r", string.Empty, StringComparison.Ordinal)
            .Split(['\n', ';', '|', ','], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (forceUpperUnderscore)
        {
            return parts
                .Select(x => x.Replace(' ', '_').ToUpperInvariant())
                .ToList();
        }

        return parts;
    }

    private bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(name);
        RefreshCommandStates();
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    private static string Compact(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return "Unknown error";
        }

        var cleaned = string.Join(" ", message.Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()));
        return cleaned.Length <= 120 ? cleaned : cleaned[..120] + "...";
    }

    private void LoadCustomerProfileIfAvailable()
    {
        if (string.IsNullOrWhiteSpace(TenantDomain))
        {
            return;
        }

        var profile = _profiles.Load(TenantDomain);
        if (profile is null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(PnPClientId))
        {
            PnPClientId = profile.ClientId;
        }
        if (string.IsNullOrWhiteSpace(PnPThumbprint))
        {
            PnPThumbprint = profile.Thumbprint;
        }
        if (string.IsNullOrWhiteSpace(ExchangeClientId))
        {
            ExchangeClientId = profile.ClientId;
        }
        if (string.IsNullOrWhiteSpace(ExchangeThumbprint))
        {
            ExchangeThumbprint = profile.Thumbprint;
        }
        if (string.IsNullOrWhiteSpace(SelectedSharePointSiteUrl)
            || string.Equals(SelectedSharePointSiteUrl, "https://hresourcing.sharepoint.com", StringComparison.OrdinalIgnoreCase)
            || string.Equals(SelectedSharePointSiteUrl, "https://", StringComparison.OrdinalIgnoreCase))
        {
            SelectedSharePointSiteUrl = profile.SiteUrl;
        }
    }

    private void SaveCustomerProfile()
    {
        if (string.IsNullOrWhiteSpace(TenantDomain) || string.IsNullOrWhiteSpace(SelectedSharePointSiteUrl))
        {
            return;
        }

        _profiles.Save(new CustomerProfile(
            TenantDomain: TenantDomain,
            SiteUrl: SelectedSharePointSiteUrl,
            ClientId: PnPClientId,
            Thumbprint: PnPThumbprint,
            UpdatedUtc: DateTimeOffset.UtcNow));
    }

    private void LoadCustomers()
    {
        Customers.Clear();
        foreach (var customer in _customerDirectory.LoadAll())
        {
            Customers.Add(customer);
        }

        if (Customers.Count > 0 && SelectedCustomer is null)
        {
            SelectedCustomer = Customers[0];
        }
    }

    private void NewCustomer()
    {
        SelectedCustomer = null;
        IsCustomerEditLocked = false;
        CustomerName = string.Empty;
        TenantDomain = string.Empty;
        SelectedSharePointSiteUrl = "https://";
        PnPClientId = string.Empty;
        PnPThumbprint = string.Empty;
        ExchangeClientId = string.Empty;
        ExchangeThumbprint = string.Empty;
        StatusMessage = "New customer profile. Enter Customer Name and Tenant Domain to enable one-time setup.";
        SignInPhaseMessage = "New customer draft ready";
        RefreshCommandStates();
    }

    private void EditCustomer()
    {
        if (SelectedCustomer is null)
        {
            return;
        }

        IsCustomerEditLocked = false;
        StatusMessage = $"Editing customer profile: {SelectedCustomer.Name}";
        SignInPhaseMessage = "Customer edit mode";
    }

    private void SaveCustomer()
    {
        if (IsCustomerEditLocked)
        {
            StatusMessage = "Profile is locked. Click Unlock Profile to edit and save.";
            return;
        }

        if (!IsUsableSharePointUrl(SelectedSharePointSiteUrl))
        {
            StatusMessage = "Enter a valid SharePoint site URL before saving customer (for example: https://tenant.sharepoint.com).";
            return;
        }

        var entry = new CustomerDirectoryEntry(
            Name: CustomerName.Trim(),
            TenantDomain: TenantDomain.Trim(),
            SiteUrl: SelectedSharePointSiteUrl.Trim(),
            PnPClientId: PnPClientId.Trim(),
            PnPThumbprint: PnPThumbprint.Trim());

        var list = Customers.ToList();
        var index = list.FindIndex(c => c.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
        if (index >= 0)
        {
            list[index] = entry;
        }
        else
        {
            list.Add(entry);
        }

        _customerDirectory.SaveAll(list);
        SaveCustomerProfile();
        LoadCustomers();
        SelectedCustomer = Customers.FirstOrDefault(c => c.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
        IsCustomerEditLocked = true;
        StatusMessage = $"Saved customer profile: {entry.Name}";
    }

    private void DeleteCustomer()
    {
        if (SelectedCustomer is null)
        {
            return;
        }

        var answer = MessageBox.Show(
            $"Delete customer profile '{SelectedCustomer.Name}'?\nThis will remove saved tenant/site/app values from the customer list.",
            "Delete Customer",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (answer != MessageBoxResult.Yes)
        {
            return;
        }

        var list = Customers.Where(c => !c.Name.Equals(SelectedCustomer.Name, StringComparison.OrdinalIgnoreCase)).ToList();
        _customerDirectory.SaveAll(list);
        LoadCustomers();

        if (Customers.Count > 0)
        {
            SelectedCustomer = Customers[0];
        }
        else
        {
            NewCustomer();
        }

        StatusMessage = "Customer profile deleted.";
    }

    private void UpsertCurrentCustomer()
    {
        if (string.IsNullOrWhiteSpace(CustomerName))
        {
            return;
        }

        var list = Customers.ToList();
        var index = list.FindIndex(c => c.Name.Equals(CustomerName, StringComparison.OrdinalIgnoreCase));
        var updated = new CustomerDirectoryEntry(
            Name: CustomerName.Trim(),
            TenantDomain: TenantDomain.Trim(),
            SiteUrl: SelectedSharePointSiteUrl.Trim(),
            PnPClientId: PnPClientId.Trim(),
            PnPThumbprint: PnPThumbprint.Trim());

        if (index >= 0)
        {
            list[index] = updated;
        }
        else
        {
            list.Add(updated);
        }

        _customerDirectory.SaveAll(list);
        LoadCustomers();
        SelectedCustomer = Customers.FirstOrDefault(c => c.Name.Equals(updated.Name, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsUsableSharePointUrl(string? url)
    {
        if (string.IsNullOrWhiteSpace(url))
        {
            return false;
        }

        var trimmed = url.Trim();
        if (string.Equals(trimmed, "https://", StringComparison.OrdinalIgnoreCase)
            || string.Equals(trimmed, "http://", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return Uri.TryCreate(trimmed, UriKind.Absolute, out var uri)
            && !string.IsNullOrWhiteSpace(uri.Host);
    }

    private static string ResolveAppVersion()
    {
        try
        {
            var processPath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(processPath))
            {
                return "unknown";
            }

            var productVersion = FileVersionInfo.GetVersionInfo(processPath).ProductVersion;
            return string.IsNullOrWhiteSpace(productVersion) ? "unknown" : productVersion;
        }
        catch
        {
            return "unknown";
        }
    }

    private static ActionableError BuildActionableError(OperationArea area, string rawDetail)
    {
        var detail = Compact(rawDetail);
        var lower = rawDetail.ToLowerInvariant();

        if (lower.Contains("consent") || lower.Contains("unauthorized") || lower.Contains("forbidden"))
        {
            return new ActionableError(
                area,
                "Permission or admin consent issue",
                $"Failed because required permissions are missing or not yet consented. Detail: {detail}",
                "Complete admin consent for the customer app, then retry.");
        }

        if (lower.Contains("certificate") || lower.Contains("thumbprint") || lower.Contains("private key"))
        {
            return new ActionableError(
                area,
                "Certificate configuration issue",
                $"Failed because the app certificate was missing, invalid, or unavailable. Detail: {detail}",
                "Re-run setup to regenerate/import certificate, then retry.");
        }

        if (lower.Contains("timeout") || lower.Contains("temporarily unavailable") || lower.Contains("rate"))
        {
            return new ActionableError(
                area,
                "Temporary service issue",
                $"Operation failed due to a temporary service/network condition. Detail: {detail}",
                "Wait a moment and retry. If it repeats, run individual steps to isolate the failing service.");
        }

        return new ActionableError(
            area,
            "Unexpected runtime issue",
            $"Operation failed with an unexpected error. Detail: {detail}",
            "Retry once. If it fails again, verify customer settings and permissions before continuing.");
    }

    private void ReportActionableSetupFailure(string rawDetail)
    {
        var error = BuildActionableError(OperationArea.SetupCustomerApp, rawDetail);
        SignInPhaseMessage = "Customer app setup failed";
        StatusMessage = $"{error.Title}. {error.UserMessage} Next: {error.NextAction}";
    }

    private string BuildLiveRunFailureDetail(string rawDetail)
    {
        var error = BuildActionableError(OperationArea.LiveRun, rawDetail);
        return $"{error.UserMessage} Next: {error.NextAction}";
    }

    private void AddPreflight(string name, bool pass, string detail)
    {
        PreflightChecks.Add(new PreflightCheckItem(pass ? "PASS" : "FAIL", name, detail, pass));
    }

    private void AddPreflightSection(string section)
    {
        PreflightChecks.Add(new PreflightCheckItem("INFO", $"--- {section} ---", string.Empty, true));
    }

    private void ToggleParserDiagnostics()
    {
        ShowParserDiagnosticsDetails = !ShowParserDiagnosticsDetails;
        if (CurrentStep == WizardStep.Review && _currentRequest is not null && PreflightChecks.Count > 0)
        {
            RunPreflight();
        }
    }

    private static string FormatDirectoryMatch(DirectoryUserMatch match)
    {
        var upn = string.IsNullOrWhiteSpace(match.UserPrincipalName) ? "(no upn)" : match.UserPrincipalName;
        var display = string.IsNullOrWhiteSpace(match.DisplayName) ? "(no display name)" : match.DisplayName;
        return $"{display} <{upn}>";
    }

    private static bool TryGetExactIdentityMatchOption(string identity, IReadOnlyList<DirectoryUserMatch> matches, out string option)
    {
        option = string.Empty;
        if (string.IsNullOrWhiteSpace(identity) || string.IsNullOrWhiteSpace(identity.Trim()) || matches.Count == 0)
        {
            return false;
        }

        var normalizedIdentity = identity.Trim();
        var exactMatches = matches
            .Where(match =>
                string.Equals(match.UserPrincipalName, normalizedIdentity, StringComparison.OrdinalIgnoreCase)
                || string.Equals(match.Mail, normalizedIdentity, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (exactMatches.Count != 1)
        {
            return false;
        }

        option = FormatDirectoryMatch(exactMatches[0]);
        return true;
    }
}

public enum WizardStep
{
    SignIn,
    Upload,
    Review,
    LiveRun,
    Summary
}

public sealed record PreflightCheckItem(string Status, string Check, string Detail, bool IsPass);
public enum OperationArea
{
    SetupCustomerApp,
    LiveRun
}

public sealed record ActionableError(OperationArea Area, string Title, string UserMessage, string NextAction);

public sealed class DirectoryUserMatchChoice : INotifyPropertyChanged
{
    private string _selectedPrincipal = string.Empty;

    public DirectoryUserMatchChoice(string sourceIdentity, IReadOnlyList<DirectoryUserMatch> matches)
    {
        SourceIdentity = sourceIdentity;
        MatchOptions = matches
            .Select(FormatMatch)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public string SourceIdentity { get; }

    public IReadOnlyList<string> MatchOptions { get; }

    public string SelectedPrincipal
    {
        get => _selectedPrincipal;
        set
        {
            if (string.Equals(_selectedPrincipal, value, StringComparison.Ordinal))
            {
                return;
            }

            _selectedPrincipal = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedPrincipal)));
        }
    }

    private static string FormatMatch(DirectoryUserMatch match)
    {
        var upn = string.IsNullOrWhiteSpace(match.UserPrincipalName) ? "(no upn)" : match.UserPrincipalName;
        var display = string.IsNullOrWhiteSpace(match.DisplayName) ? "(no display name)" : match.DisplayName;
        return $"{display} <{upn}>";
    }
}
