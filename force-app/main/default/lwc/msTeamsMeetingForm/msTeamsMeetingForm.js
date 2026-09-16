import { LightningElement, api } from 'lwc';
import createMeeting from '@salesforce/apex/MSTeamsMeetingController.createMeeting';
import searchAttendees from '@salesforce/apex/MSTeamsMeetingController.searchAttendees';
import TEAMS_LOGO from '@salesforce/resourceUrl/MS_Teams_Logo';

const pad = (value) => String(value).padStart(2, '0');

export default class MsTeamsMeetingForm extends LightningElement {
    @api recordId;
    @api context;
    @api testMode = false;
    teamsLogoUrl = TEAMS_LOGO;
    step = 1;
    saving = false;
    result;
    searchResults = [];
    selectedAttendees = [];
    attendeeTab = 'salesforce';
    searchTimer;
    errorMessage;
    form = {
        subject: '', description: '', meetingType: 'Customer Meeting', date: '', startTime: '15:00',
        duration: '60', microsoftTimeZone: 'India Standard Time', reminderMinutesBeforeStart: '15',
        allowNewTimeProposals: true, personalMessage: '', allowMeetingChat: true,
        allowScreenSharing: true, allowRecording: true, allowTranscription: true, recordAutomatically: true, enableTranscription: true,
        lobbyRequired: false, meetingAccess: 'Invited People', teamsChannelId: '', location: ''
    };

    connectedCallback() {
        const now = new Date();
        const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
        this.form = {
            ...this.form,
            subject: this.context?.defaultSubject || `${this.context?.recordName || 'Customer'} meeting`,
            description: this.context?.defaultDescription || '',
            date: `${tomorrow.getFullYear()}-${pad(tomorrow.getMonth() + 1)}-${pad(tomorrow.getDate())}`
        };
        this.selectedAttendees = (this.context?.defaultRecipientEmails || []).map((email) => ({
            key: email, displayName: email, email, title: 'Record email', initials: this.initials(email)
        }));
    }

    get isStep1() { return this.step === 1; }
    get isStep2() { return this.step === 2; }
    get isStep3() { return this.step === 3; }
    get isStep4() { return this.step === 4; }
    get isSuccess() { return Boolean(this.result); }
    get modalTitle() { return this.isSuccess ? 'Meeting Scheduled' : 'Schedule Microsoft Teams Meeting'; }
    get nextLabel() { return this.step === 1 ? 'Next: Add Attendees' : this.step === 2 ? 'Next: Teams Options' : 'Next: Review & Schedule'; }
    get showBack() { return this.step > 1; }
    get hasSearchResults() { return this.searchResults.length > 0; }
    get selectedCount() { return this.selectedAttendees.length; }
    get recordName() { return this.context?.recordName || 'Current record'; }
    get objectLabel() { return (this.context?.objectApiName || 'Record').replace(/__c$/, '').replaceAll('_', ' '); }
    get dateTimeSummary() {
        if (!this.form.date || !this.form.startTime) return '';
        const date = new Date(`${this.form.date}T${this.form.startTime}:00`);
        return `${new Intl.DateTimeFormat('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' }).format(date)} · ${new Intl.DateTimeFormat('en-US', { hour: 'numeric', minute: '2-digit' }).format(date)} · ${this.form.duration} min`;
    }
    get step1Class() { return this.stepClass(1); }
    get step2Class() { return this.stepClass(2); }
    get step3Class() { return this.stepClass(3); }
    get step4Class() { return this.stepClass(4); }
    get sfTabClass() { return `attendee-tab${this.attendeeTab === 'salesforce' ? ' active' : ''}`; }
    get emailTabClass() { return `attendee-tab${this.attendeeTab === 'email' ? ' active' : ''}`; }
    get suggestedTabClass() { return `attendee-tab${this.attendeeTab === 'suggested' ? ' active' : ''}`; }
    get meetingTypeOptions() { return [{label:'Customer Meeting',value:'Customer Meeting'},{label:'Internal Meeting',value:'Internal Meeting'},{label:'Product Demo',value:'Product Demo'},{label:'Contract Review',value:'Contract Review'}]; }
    get durationOptions() { return [15,30,45,60,90,120].map((v) => ({label:`${v} minutes`,value:String(v)})); }
    get timeZoneOptions() { return [{label:'(GMT+05:30) Chennai, Kolkata, Mumbai, New Delhi',value:'India Standard Time'},{label:'(GMT-08:00) Pacific Time',value:'Pacific Standard Time'},{label:'(GMT-05:00) Eastern Time',value:'Eastern Standard Time'},{label:'(GMT+00:00) London',value:'GMT Standard Time'},{label:'(GMT+01:00) Central Europe',value:'W. Europe Standard Time'}]; }
    get reminderOptions() { return [{label:'No reminder',value:'-1'},{label:'5 minutes before',value:'5'},{label:'15 minutes before',value:'15'},{label:'30 minutes before',value:'30'},{label:'1 hour before',value:'60'}]; }
    get accessOptions() { return [{label:'Only invited people can join (Recommended)',value:'Invited People'},{label:'People in my organization can join',value:'Organization'},{label:'Anyone with the link can join',value:'Anyone'},{label:'My organization and trusted organizations',value:'Organization and Trusted'}]; }

    get accessChoices() {
        return this.accessOptions.map((option) => ({ ...option, checked: this.form.meetingAccess === option.value }));
    }

    stepClass(number) {
        return `step ${this.step === number ? 'current' : ''} ${this.step > number || this.result ? 'complete' : ''}`;
    }
    handleField(event) {
        const field = event.target.dataset.field;
        this.form = { ...this.form, [field]: event.detail?.value ?? event.target.value };
        this.errorMessage = undefined;
    }
    handleToggle(event) {
        const field = event.target.dataset.field;
        const checked = event.target.checked;
        const changes = { [field]: checked };
        if (field === 'allowRecording' && !checked) changes.recordAutomatically = false;
        if (field === 'recordAutomatically' && checked) changes.allowRecording = true;
        if (field === 'allowTranscription' && !checked) changes.enableTranscription = false;
        if (field === 'enableTranscription' && checked) changes.allowTranscription = true;
        this.form = { ...this.form, ...changes };
    }
    close() { this.dispatchEvent(new CustomEvent('close')); }
    back() { this.step = Math.max(1, this.step - 1); this.errorMessage = undefined; }
    next() {
        if (!this.validateStep()) return;
        this.step += 1;
    }
    validateStep() {
        this.errorMessage = undefined;
        if (this.step === 1) {
            const inputs = [...this.template.querySelectorAll('[data-required="true"]')];
            const valid = inputs.reduce((ok, input) => { input.reportValidity(); return input.checkValidity() && ok; }, true);
            if (!valid) return false;
        }
        if (this.step === 2 && !this.selectedAttendees.length) {
            this.errorMessage = 'Select or add at least one attendee.';
            return false;
        }
        return true;
    }
    chooseTab(event) { this.attendeeTab = event.currentTarget.dataset.tab; }
    handleSearch(event) {
        window.clearTimeout(this.searchTimer);
        const term = event.target.value;
        if (!term || term.trim().length < 2) { this.searchResults = []; return; }
        this.searchTimer = window.setTimeout(async () => {
            try {
                const rows = await searchAttendees({ recordId: this.recordId, searchTerm: term });
                const selected = new Set(this.selectedAttendees.map((item) => item.email.toLowerCase()));
                this.searchResults = rows.filter((row) => !selected.has(row.email.toLowerCase())).map((row) => ({ ...row, key: row.recordId, initials: this.initials(row.displayName) }));
            } catch (error) { this.errorMessage = this.reduceError(error); }
        }, 300);
    }
    addSearchResult(event) {
        const row = this.searchResults.find((item) => item.key === event.currentTarget.dataset.id);
        if (!row) return;
        this.selectedAttendees = [...this.selectedAttendees, row];
        this.searchResults = this.searchResults.filter((item) => item.key !== row.key);
    }
    addEmail() {
        const input = this.template.querySelector('[data-id="manual-email"]');
        input.reportValidity();
        if (!input.checkValidity()) return;
        const email = input.value.trim().toLowerCase();
        if (!this.selectedAttendees.some((item) => item.email.toLowerCase() === email)) {
            this.selectedAttendees = [...this.selectedAttendees, { key: email, displayName: email, email, title: 'Email attendee', initials: this.initials(email) }];
        }
        input.value = '';
    }
    removeAttendee(event) {
        this.selectedAttendees = this.selectedAttendees.filter((item) => item.key !== event.currentTarget.dataset.id);
    }
    async schedule() {
        if (this.testMode) {
            this.errorMessage = 'This is a UI test preview. Disable test mode and connect Microsoft 365 to schedule a real meeting.';
            return;
        }
        this.saving = true;
        this.errorMessage = undefined;
        try {
            const start = `${this.form.date}T${this.form.startTime}:00`;
            const [year, month, day] = this.form.date.split('-').map(Number);
            const [hour, minute] = this.form.startTime.split(':').map(Number);
            const endDate = new Date(Date.UTC(year, month - 1, day, hour, minute + Number(this.form.duration)));
            const end = `${endDate.getUTCFullYear()}-${pad(endDate.getUTCMonth()+1)}-${pad(endDate.getUTCDate())}T${pad(endDate.getUTCHours())}:${pad(endDate.getUTCMinutes())}:00`;
            const description = [this.form.description, this.form.personalMessage].filter(Boolean).join('\n\n');
            this.result = await createMeeting({ request: {
                recordId: this.recordId, subject: this.form.subject, description,
                attendeeEmails: this.selectedAttendees.map((item) => item.email), startDateTime: start,
                endDateTime: end, microsoftTimeZone: this.form.microsoftTimeZone,
                recordAutomatically: this.form.recordAutomatically, enableTranscription: this.form.enableTranscription,
                allowRecording: this.form.allowRecording, allowTranscription: this.form.allowTranscription,
                meetingType: this.form.meetingType, allowMeetingChat: this.form.allowMeetingChat,
                allowScreenSharing: this.form.allowScreenSharing, lobbyRequired: this.form.lobbyRequired,
                meetingAccess: this.form.meetingAccess, teamsChannelId: this.form.teamsChannelId,
                location: this.form.location, reminderMinutesBeforeStart: Number(this.form.reminderMinutesBeforeStart),
                allowNewTimeProposals: this.form.allowNewTimeProposals
            }});
            this.dispatchEvent(new CustomEvent('scheduled', { detail: this.result }));
        } catch (error) { this.errorMessage = this.reduceError(error); }
        finally { this.saving = false; }
    }
    scheduleAnother() { this.result = undefined; this.step = 1; }
    joinMeeting() { if (this.result?.joinUrl) window.open(this.result.joinUrl, '_blank', 'noopener'); }
    initials(value) { return (value || '?').split(/[@ ._-]/).filter(Boolean).slice(0,2).map((part) => part[0].toUpperCase()).join(''); }
    reduceError(error) { return error?.body?.message || error?.message || 'The meeting could not be scheduled.'; }
}
