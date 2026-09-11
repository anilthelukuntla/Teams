import { LightningElement, api, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import { CurrentPageReference } from 'lightning/navigation';
import getContext from '@salesforce/apex/MSTeamsMeetingController.getContext';
import getMeetings from '@salesforce/apex/MSTeamsMeetingController.getMeetings';
import getMicrosoftConnectionStatus from '@salesforce/apex/MSTeamsMeetingController.getMicrosoftConnectionStatus';
import getMicrosoftAuthorizationUrl from '@salesforce/apex/MSTeamsMeetingController.getMicrosoftAuthorizationUrl';
import TEAMS_LOGO from '@salesforce/resourceUrl/MS_Teams_Logo';

const STATUS_OPTIONS = [
    { label: 'All Status', value: 'all' },
    { label: 'Scheduled', value: 'Scheduled' },
    { label: 'In Progress', value: 'In Progress' },
    { label: 'Completed', value: 'Completed' },
    { label: 'Cancelled', value: 'Cancelled' },
    { label: 'Deleted', value: 'Deleted' },
    { label: 'Sync Failed', value: 'Sync Failed' }
];

export default class MsTeamsMeetings extends LightningElement {
    _recordId;
    componentConnected = false;
    lastLoadedRecordId;

    @api
    get recordId() {
        return this._recordId;
    }

    set recordId(value) {
        this.setRecordId(value);
    }

    @wire(CurrentPageReference)
    resolvePageReference(pageReference) {
        const pageRecordId =
            pageReference?.attributes?.recordId ||
            pageReference?.state?.recordId ||
            pageReference?.state?.c__recordId;
        this.setRecordId(pageRecordId);
    }

    teamsLogoUrl = TEAMS_LOGO;
    isLoading = true;
    isConnecting = false;
    hasError = false;
    errorMessage;
    connectionMessage;
    connection;
    context;
    meetings = [];
    activeTab = 'upcoming';
    searchTerm = '';
    statusFilter = 'all';
    sortOrder = 'dateAsc';
    calendarCursor = new Date();
    connectionPoll;
    showScheduleModal = false;
    selectedMeetingId;
    selectedDetailsTab = 'overview';

    statusOptions = STATUS_OPTIONS;
    sortOptions = [
        { label: 'Sort by Date', value: 'dateAsc' },
        { label: 'Newest First', value: 'dateDesc' },
        { label: 'Title A–Z', value: 'titleAsc' }
    ];

    connectedCallback() {
        this.componentConnected = true;
        if (this.recordId) {
            this.loadWorkspace();
        } else {
            this.isLoading = false;
        }
    }

    disconnectedCallback() {
        this.componentConnected = false;
        this.stopConnectionPolling();
    }

    @api
    async refresh() {
        await this.loadWorkspace();
    }

    async loadWorkspace() {
        if (!this.recordId) {
            this.hasError = false;
            this.isLoading = false;
            return;
        }
        const requestedRecordId = this.recordId;
        this.lastLoadedRecordId = requestedRecordId;
        this.isLoading = true;
        this.hasError = false;
        const [contextResult, meetingsResult, connectionResult] =
            await Promise.allSettled([
                getContext({ recordId: requestedRecordId }),
                getMeetings({ recordId: requestedRecordId }),
                getMicrosoftConnectionStatus()
            ]);

        if (contextResult.status === 'rejected' || meetingsResult.status === 'rejected') {
            this.hasError = true;
            this.errorMessage = this.reduceError(
                contextResult.status === 'rejected' ? contextResult.reason : meetingsResult.reason
            );
        } else {
            this.context = contextResult.value;
            this.meetings = (meetingsResult.value || []).map((meeting) => this.mapMeeting(meeting));
        }

        if (connectionResult.status === 'fulfilled') {
            this.connection = connectionResult.value;
            this.connectionMessage = connectionResult.value?.message;
        } else {
            this.connection = { connected: false };
            this.connectionMessage = this.reduceError(connectionResult.reason);
        }
        this.isLoading = false;
    }

    setRecordId(value) {
        if (!value || value === this._recordId) {
            return;
        }
        this._recordId = value;
        if (this.componentConnected && value !== this.lastLoadedRecordId) {
            this.loadWorkspace();
        }
    }

    get isDisconnected() {
        return this.connection?.connected !== true;
    }

    get isMissingRecordContext() {
        return !this.recordId;
    }

    get connectionEmail() {
        return this.connection?.email || this.connection?.userPrincipalName || '';
    }

    get recordName() {
        return this.context?.recordName || 'Current record';
    }

    get objectLabel() {
        return (this.context?.objectApiName || 'Salesforce Record').replace(/__c$/, '').replaceAll('_', ' ');
    }

    get upcomingMeetings() {
        return this.meetings.filter((meeting) => meeting.isUpcoming);
    }

    get pastMeetings() {
        return this.meetings.filter((meeting) => !meeting.isUpcoming);
    }

    get upcomingCount() {
        return this.upcomingMeetings.length;
    }

    get completedCount() {
        return this.meetings.filter((meeting) => meeting.status === 'Completed').length;
    }

    get recordedCount() {
        return this.meetings.filter((meeting) => meeting.recordingRequested).length;
    }

    get visibleMeetings() {
        let result;
        if (this.activeTab === 'past') {
            result = this.pastMeetings;
        } else if (this.activeTab === 'recordings') {
            result = this.meetings.filter((meeting) => meeting.recordingRequested);
        } else {
            result = this.upcomingMeetings;
        }
        if (this.statusFilter !== 'all') {
            result = result.filter((meeting) => meeting.status === this.statusFilter);
        }
        if (this.searchTerm) {
            result = result.filter((meeting) =>
                `${meeting.subject} ${meeting.description}`.toLowerCase().includes(this.searchTerm)
            );
        }
        return [...result].sort((left, right) => {
            if (this.sortOrder === 'titleAsc') {
                return left.subject.localeCompare(right.subject);
            }
            return this.sortOrder === 'dateDesc'
                ? right.sortTimestamp - left.sortTimestamp
                : left.sortTimestamp - right.sortTimestamp;
        });
    }

    get hasVisibleMeetings() {
        return this.visibleMeetings.length > 0;
    }

    get emptyTitle() {
        return this.meetings.length ? 'No matching meetings' : 'No meetings yet';
    }

    get emptyMessage() {
        return this.meetings.length
            ? 'Try changing the search, status, or meeting view.'
            : `Schedule a Microsoft Teams meeting for ${this.recordName} to get started.`;
    }

    get showEmptyScheduleAction() {
        return this.meetings.length === 0;
    }

    get upcomingTabClass() {
        return this.tabClass('upcoming');
    }

    get pastTabClass() {
        return this.tabClass('past');
    }

    get recordingsTabClass() {
        return this.tabClass('recordings');
    }

    get calendarMonthLabel() {
        return new Intl.DateTimeFormat('en-US', { month: 'long', year: 'numeric' }).format(this.calendarCursor);
    }

    get calendarDays() {
        const year = this.calendarCursor.getFullYear();
        const month = this.calendarCursor.getMonth();
        const firstWeekday = new Date(year, month, 1).getDay();
        const dayCount = new Date(year, month + 1, 0).getDate();
        const today = new Date();
        const meetingDays = new Set(
            this.meetings
                .filter((meeting) => meeting.dateParts?.year === year && meeting.dateParts?.month === month)
                .map((meeting) => meeting.dateParts.day)
        );
        const days = [];
        for (let index = 0; index < firstWeekday; index += 1) {
            days.push({ key: `blank-${index}`, label: '', className: 'calendar-day blank' });
        }
        for (let day = 1; day <= dayCount; day += 1) {
            const isToday = today.getFullYear() === year && today.getMonth() === month && today.getDate() === day;
            const hasMeeting = meetingDays.has(day);
            days.push({
                key: `day-${day}`,
                label: day,
                className: `calendar-day${isToday ? ' today' : ''}${hasMeeting ? ' has-meeting' : ''}`
            });
        }
        return days;
    }

    get calendarMeetings() {
        return this.upcomingMeetings.slice(0, 2);
    }

    async handleConnect() {
        this.isConnecting = true;
        this.connectionMessage = 'Opening Microsoft sign-in…';
        const authWindow = window.open('', '_blank', 'noopener');
        try {
            const authorizationUrl = await getMicrosoftAuthorizationUrl();
            if (authWindow) {
                authWindow.location = authorizationUrl;
            } else {
                window.location.assign(authorizationUrl);
            }
            this.startConnectionPolling();
        } catch (error) {
            if (authWindow) {
                authWindow.close();
            }
            this.connectionMessage = this.reduceError(error);
            this.isConnecting = false;
        }
    }

    startConnectionPolling() {
        this.stopConnectionPolling();
        let attempts = 0;
        this.connectionPoll = window.setInterval(async () => {
            attempts += 1;
            try {
                const result = await getMicrosoftConnectionStatus();
                if (result?.connected) {
                    this.connection = result;
                    this.connectionMessage = result.message;
                    this.isConnecting = false;
                    this.stopConnectionPolling();
                    this.showToast('Microsoft 365 connected', `Connected as ${result.email || result.userPrincipalName}.`, 'success');
                }
            } catch (error) {
                this.connectionMessage = this.reduceError(error);
            }
            if (attempts >= 24) {
                this.isConnecting = false;
                this.stopConnectionPolling();
            }
        }, 5000);
    }

    stopConnectionPolling() {
        if (this.connectionPoll) {
            window.clearInterval(this.connectionPoll);
            this.connectionPoll = undefined;
        }
    }

    handleTab(event) {
        this.activeTab = event.currentTarget.dataset.tab;
    }

    handleSearch(event) {
        this.searchTerm = (event.target.value || '').trim().toLowerCase();
    }

    handleStatusFilter(event) {
        this.statusFilter = event.detail.value;
    }

    handleSort(event) {
        this.sortOrder = event.detail.value;
    }

    handleSchedule() {
        this.showScheduleModal = true;
    }

    handleScheduleClose() {
        this.showScheduleModal = false;
    }

    async handleMeetingScheduled() {
        await this.loadWorkspace();
    }

    async handleSync() {
        await this.loadWorkspace();
        if (!this.hasError) {
            this.showToast('Meetings refreshed', 'The latest Salesforce meeting data is displayed.', 'success');
        }
    }

    handleOpenOutlook() {
        window.open('https://outlook.office.com/calendar/view/month', '_blank', 'noopener');
    }

    handleSettings() {
        this.dispatchEvent(new CustomEvent('meetingsettings', { bubbles: true, composed: true }));
    }

    handleJoin(event) {
        const url = event.currentTarget.dataset.url;
        if (url) {
            window.open(url, '_blank', 'noopener');
        }
    }

    handleDetails(event) {
        this.selectedDetailsTab = 'overview';
        this.selectedMeetingId = event.currentTarget.dataset.id;
    }

    handleDetailsClose() {
        this.selectedMeetingId = undefined;
    }

    handleDetailsNavigate(event) {
        this.dispatchMeetingAction(event.detail.action, event.detail.meetingId);
    }

    handleDetailsAction(event) {
        this.selectedMeetingId = undefined;
        this.dispatchMeetingAction(event.detail.action, event.detail.meetingId);
    }

    handleMeetingMenu(event) {
        if (event.detail.value === 'participants') {
            this.selectedDetailsTab = 'participants';
            this.selectedMeetingId = event.currentTarget.dataset.id;
            return;
        }
        this.dispatchMeetingAction(event.detail.value, event.currentTarget.dataset.id);
    }

    handleViewAll() {
        this.activeTab = 'upcoming';
    }

    previousMonth() {
        this.calendarCursor = new Date(this.calendarCursor.getFullYear(), this.calendarCursor.getMonth() - 1, 1);
    }

    nextMonth() {
        this.calendarCursor = new Date(this.calendarCursor.getFullYear(), this.calendarCursor.getMonth() + 1, 1);
    }

    dispatchMeetingAction(action, meetingId) {
        this.dispatchEvent(new CustomEvent('meetingaction', {
            detail: { action, meetingId },
            bubbles: true,
            composed: true
        }));
    }

    tabClass(tabName) {
        return `tab-button${this.activeTab === tabName ? ' active' : ''}`;
    }

    mapMeeting(meeting) {
        const start = this.parseLocalDateTime(meeting.startDateTime);
        const end = this.parseLocalDateTime(meeting.endDateTime);
        const terminalStatuses = new Set(['Completed', 'Cancelled', 'Deleted']);
        const isUpcoming = !terminalStatuses.has(meeting.status) && (!end || end.getTime() >= Date.now());
        const attendeeEmails = (meeting.attendeeEmails || '').split(/[;,]/).map((email) => email.trim()).filter(Boolean);
        const attendees = attendeeEmails.slice(0, 3).map((email) => ({ email, initials: this.initialsFromEmail(email) }));
        const description = this.plainText(meeting.description) || 'No agenda provided.';
        const duration = start && end ? Math.max(0, Math.round((end - start) / 60000)) : null;
        const statusToken = (meeting.status || 'Scheduled').toLowerCase().replaceAll(' ', '-');
        return {
            ...meeting,
            description,
            attendees,
            extraAttendeeCount: Math.max(0, attendeeEmails.length - attendees.length),
            attendeeAriaLabel: `${attendeeEmails.length} attendees`,
            isUpcoming,
            canJoin: isUpcoming && Boolean(meeting.joinUrl),
            sortTimestamp: start?.getTime() || 0,
            dateParts: start ? { year: start.getFullYear(), month: start.getMonth(), day: start.getDate() } : null,
            dateLabel: start ? new Intl.DateTimeFormat('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' }).format(start) : 'Date unavailable',
            timeLabel: start ? `${this.formatTime(start)}${end ? ` – ${this.formatTime(end)}` : ''}${duration !== null ? ` (${duration} min)` : ''}` : 'Time unavailable',
            shortTimeLabel: start ? `${this.formatTime(start)}${end ? ` – ${this.formatTime(end)}` : ''}` : '',
            statusClass: `status-pill status-${statusToken}`,
            cardClass: `meeting-card meeting-${statusToken}`,
            recordingLabel: meeting.recordingRequested ? (isUpcoming ? 'Recording ON' : 'Recording Requested') : 'Recording OFF',
            transcriptLabel: meeting.transcriptionRequested ? (isUpcoming ? 'Transcript ON' : 'Transcript Requested') : 'Transcript OFF',
            recordingClass: `feature-pill ${meeting.recordingRequested ? 'enabled' : 'disabled'}`,
            transcriptClass: `feature-pill ${meeting.transcriptionRequested ? 'enabled' : 'disabled'}`
        };
    }

    parseLocalDateTime(value) {
        if (!value) return null;
        const match = value.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
        if (!match) return null;
        return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4]), Number(match[5]), Number(match[6] || 0));
    }

    formatTime(date) {
        return new Intl.DateTimeFormat('en-US', { hour: 'numeric', minute: '2-digit' }).format(date);
    }

    initialsFromEmail(email) {
        const local = email.split('@')[0] || '';
        return local.split(/[._-]/).filter(Boolean).slice(0, 2).map((part) => part[0].toUpperCase()).join('') || '?';
    }

    plainText(html) {
        return (html || '').replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim();
    }

    reduceError(error) {
        return error?.body?.message || error?.message || 'Something went wrong while loading Microsoft Teams meetings.';
    }

    showToast(title, message, variant) {
        this.dispatchEvent(new ShowToastEvent({ title, message, variant }));
    }
}
