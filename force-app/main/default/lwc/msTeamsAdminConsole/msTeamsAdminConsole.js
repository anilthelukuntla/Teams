import { LightningElement } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getSetupInfo from '@salesforce/apex/MSTeamsAdminController.getSetupInfo';
import saveSettings from '@salesforce/apex/MSTeamsAdminController.saveSettings';
import getConnectionStatus from '@salesforce/apex/MSTeamsMeetingController.getMicrosoftConnectionStatus';
import getAuthorizationUrl from '@salesforce/apex/MSTeamsMeetingController.getMicrosoftAuthorizationUrl';
import TEAMS_LOGO from '@salesforce/resourceUrl/MS_Teams_Logo';

export default class MsTeamsAdminConsole extends LightningElement {
    teamsLogoUrl = TEAMS_LOGO;
    info = {};
    connection = {};
    isLoading = true;
    isSaving = false;
    activeSection = 'connection';
    authorizationUrl;

    connectedCallback() { this.load(); }

    async load() {
        this.isLoading = true;
        const [setupResult, connectionResult] = await Promise.allSettled([getSetupInfo(), getConnectionStatus()]);
        if (setupResult.status === 'fulfilled') this.info = { ...setupResult.value };
        else this.toast('Unable to load setup', this.errorMessage(setupResult.reason), 'error');
        if (connectionResult.status === 'fulfilled') this.connection = connectionResult.value || {};
        else this.connection = { connected: false, message: this.errorMessage(connectionResult.reason) };
        if (!this.connection.connected) {
            try { this.authorizationUrl = await getAuthorizationUrl(); } catch (error) { this.authorizationUrl = undefined; }
        }
        this.isLoading = false;
    }

    get isConnectionSection() { return this.activeSection === 'connection'; }
    get isSyncSection() { return this.activeSection === 'sync'; }
    get connectionLabel() { return this.connection.connected ? 'Connected' : 'Disconnected'; }
    get connectionClass() { return `status-pill ${this.connection.connected ? 'healthy' : 'warning'}`; }
    get connectedIdentity() { return this.connection.email || this.connection.userPrincipalName || 'Current Salesforce user'; }
    get formattedLastSync() {
        return this.info.lastSync ? new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(this.info.lastSync)) : 'Not synchronized yet';
    }
    get connectionNavClass() { return `nav-item ${this.isConnectionSection ? 'active' : ''}`; }
    get syncNavClass() { return `nav-item ${this.isSyncSection ? 'active' : ''}`; }

    selectSection(event) {
        this.activeSection = event.currentTarget.dataset.section === 'sync' ? 'sync' : 'connection';
    }
    handleToggle(event) { this.info = { ...this.info, [event.target.name]: event.target.checked }; }
    handleWindow(event) { this.info = { ...this.info, syncWindowDays: Number(event.detail.value) }; }

    connect() {
        if (!this.authorizationUrl) {
            this.toast('Microsoft connection', 'The authorization URL is not ready. Refresh and try again.', 'warning');
            return;
        }
        window.open(this.authorizationUrl, '_blank', 'noopener,noreferrer');
    }

    async save() {
        this.isSaving = true;
        try {
            this.info = await saveSettings({ input: this.info });
            this.toast('Settings saved', 'Microsoft Teams synchronization settings were updated.', 'success');
        } catch (error) {
            this.toast('Unable to save settings', this.errorMessage(error), 'error');
        } finally { this.isSaving = false; }
    }

    toast(title, message, variant) { this.dispatchEvent(new ShowToastEvent({ title, message, variant })); }
    errorMessage(error) { return error?.body?.message || error?.message || 'An unexpected error occurred.'; }
}
