import { LightningElement, api } from 'lwc';
import getMeetingDetails from '@salesforce/apex/MSTeamsMeetingController.getMeetingDetails';
import TEAMS_LOGO from '@salesforce/resourceUrl/MS_Teams_Logo';

export default class MsTeamsMeetingDetails extends LightningElement {
    @api recordId;
    @api meetingId;
    @api context;
    teamsLogoUrl = TEAMS_LOGO;
    loading = true;
    errorMessage;
    details;

    connectedCallback() { this.load(); }
    async load() {
        this.loading = true;
        this.errorMessage = undefined;
        try {
            const result = await getMeetingDetails({ recordId: this.recordId, scheduleId: this.meetingId });
            this.details = this.mapDetails(result);
        } catch (error) { this.errorMessage = this.reduceError(error); }
        finally { this.loading = false; }
    }
    get meeting() { return this.details?.meeting || {}; }
    get hasError() { return Boolean(this.errorMessage); }
    get canJoin() { return Boolean(this.meeting.joinUrl) && !['Cancelled','Deleted','Completed'].includes(this.meeting.status); }
    get hasWebLink() { return Boolean(this.meeting.webLink); }
    get participantCount() { return this.details?.participantCount || 0; }
    get attendedCount() { return this.details?.attendedCount || 0; }
    get durationLabel() { return this.details?.durationLabel || '—'; }
    get recordingLabel() { return this.details?.recordingLabel || 'Not available'; }
    get transcriptLabel() { return this.details?.transcriptLabel || 'Not available'; }
    get description() { return this.details?.description || 'No agenda was provided.'; }
    get organizerEmail() { return this.meeting.organizerEmail || 'Organizer unavailable'; }
    get meetingType() { return this.meeting.meetingType || 'Customer Meeting'; }
    get location() { return this.details?.location || 'Microsoft Teams (Online)'; }
    get relatedRecords() { return this.details?.relatedRecords || []; }
    get statusClass() { return `status status-${(this.meeting.status || 'scheduled').toLowerCase().replaceAll(' ','-')}`; }
    get dateLabel() { return this.details?.dateLabel || 'Date unavailable'; }
    get timeLabel() { return this.details?.timeLabel || 'Time unavailable'; }
    get syncError() { return this.meeting.lastGraphError; }
    get chatLabel() { return this.details?.allowMeetingChat ? 'Enabled' : 'Disabled'; }
    get sharingLabel() { return this.details?.allowScreenSharing ? 'Enabled' : 'Disabled'; }
    get lobbyLabel() { return this.details?.lobbyRequired ? 'Required' : 'Not required'; }
    get hasRelatedRecords() { return this.relatedRecords.length > 0; }

    mapDetails(result) {
        const meeting = result?.meeting || {};
        const start = this.parseLocal(meeting.startDateTime);
        const end = this.parseLocal(meeting.endDateTime);
        const emailCount = (meeting.attendeeEmails || '').split(/[;,]/).map((v) => v.trim()).filter(Boolean).length;
        const participants = result?.participants || [];
        const artifacts = result?.artifacts || [];
        const recording = artifacts.find((item) => item.artifactType === 'Recording');
        const transcript = artifacts.find((item) => item.artifactType === 'Transcript');
        const seconds = Number(result?.actualDurationSeconds || 0);
        const scheduledMinutes = start && end ? Math.max(0, Math.round((end - start) / 60000)) : 0;
        return {
            ...result,
            description: this.plainText(meeting.description),
            participantCount: Math.max(emailCount, participants.length),
            attendedCount: participants.filter((person) => person.actuallyJoined).length,
            durationLabel: seconds ? `${Math.round(seconds / 60)} min` : scheduledMinutes ? `${scheduledMinutes} min` : '—',
            dateLabel: start ? new Intl.DateTimeFormat('en-US',{weekday:'short',month:'short',day:'numeric',year:'numeric'}).format(start) : 'Date unavailable',
            timeLabel: start ? `${this.time(start)}${end ? ` – ${this.time(end)}` : ''}` : 'Time unavailable',
            recordingLabel: recording?.processingStatus || (meeting.recordingRequested ? 'Requested' : 'Not available'),
            transcriptLabel: transcript?.processingStatus || (meeting.transcriptionRequested ? 'Requested' : 'Not available'),
            relatedRecords: (result?.relatedRecords || []).map((record) => ({...record, key:`${record.objectApiName}-${record.recordId}`, label:record.recordLabel || record.recordId, objectLabel:(record.objectApiName || 'Record').replace(/__c$/,'').replaceAll('_',' ')}))
        };
    }
    close() { this.dispatchEvent(new CustomEvent('close')); }
    join() { if (this.meeting.joinUrl) window.open(this.meeting.joinUrl,'_blank','noopener'); }
    openOutlook() { if (this.meeting.webLink) window.open(this.meeting.webLink,'_blank','noopener'); }
    chooseTab(event) {
        const action = event.currentTarget.dataset.tab;
        if (action === 'overview') return;
        this.dispatchEvent(new CustomEvent('navigate', { detail: { action, meetingId: this.meetingId } }));
    }
    moreAction(event) { this.dispatchEvent(new CustomEvent('action', { detail: { action:event.detail.value, meetingId:this.meetingId } })); }
    parseLocal(value) {
        const match = value?.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
        return match ? new Date(+match[1],+match[2]-1,+match[3],+match[4],+match[5],+(match[6]||0)) : null;
    }
    time(date) { return new Intl.DateTimeFormat('en-US',{hour:'numeric',minute:'2-digit'}).format(date); }
    plainText(html) { return (html || '').replace(/<br\s*\/?>/gi,' ').replace(/<[^>]+>/g,'').replace(/&nbsp;/g,' ').trim(); }
    reduceError(error) { return error?.body?.message || error?.message || 'The meeting details could not be loaded.'; }
}
