import { LightningElement, api } from 'lwc';
import getMeetingDetails from '@salesforce/apex/MSTeamsMeetingController.getMeetingDetails';
import getTranscriptContent from '@salesforce/apex/MSTeamsMeetingController.getTranscriptContent';
import TEAMS_LOGO from '@salesforce/resourceUrl/MS_Teams_Logo';

export default class MsTeamsMeetingDetails extends LightningElement {
    @api recordId;
    @api meetingId;
    @api context;
    @api initialTab = 'overview';
    teamsLogoUrl = TEAMS_LOGO;
    loading = true;
    errorMessage;
    details;
    activeTab = 'overview';
    participantSearch = '';
    participantFilter = 'all';
    transcriptLoading = false;
    transcriptLoaded = false;
    transcriptError;
    transcriptSegments = [];
    transcriptSearch = '';

    connectedCallback() { this.activeTab = this.initialTab || 'overview'; this.load(); }
    async load() {
        this.loading = true;
        this.errorMessage = undefined;
        try {
            const result = await getMeetingDetails({ recordId: this.recordId, scheduleId: this.meetingId });
            this.details = this.mapDetails(result);
            if (this.activeTab === 'transcript' && this.transcriptAvailable) await this.loadTranscript();
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
    get isOverview() { return this.activeTab === 'overview'; }
    get isParticipants() { return this.activeTab === 'participants'; }
    get isArtifacts() { return this.activeTab === 'transcript' || this.activeTab === 'recording'; }
    get overviewTabClass() { return this.tabClass('overview'); }
    get participantsTabClass() { return this.tabClass('participants'); }
    get transcriptTabClass() { return this.tabClass('transcript'); }
    get recordingTabClass() { return this.tabClass('recording'); }
    get isTranscriptView() { return this.activeTab === 'transcript'; }
    get isRecordingView() { return this.activeTab === 'recording'; }
    get recordingArtifact() { return (this.details?.artifacts || []).find((item) => item.artifactType === 'Recording'); }
    get transcriptArtifact() { return (this.details?.artifacts || []).find((item) => item.artifactType === 'Transcript'); }
    get recordingAvailable() { return this.recordingArtifact?.processingStatus?.toLowerCase() === 'available'; }
    get transcriptAvailable() { return Boolean(this.transcriptArtifact?.contentUrl); }
    get artifactStateTitle() { return this.activeTab === 'recording' ? 'Recording is not available yet' : 'Transcript is not available yet'; }
    get artifactStateMessage() { return this.activeTab === 'recording' ? 'Microsoft is still processing the recording, or recording was not enabled for this meeting.' : 'Microsoft is still generating the transcript, or transcription was not enabled.'; }
    get filteredTranscriptSegments() { return !this.transcriptSearch ? this.transcriptSegments : this.transcriptSegments.filter((segment) => `${segment.speaker} ${segment.text}`.toLowerCase().includes(this.transcriptSearch)); }
    get hasTranscriptSegments() { return this.filteredTranscriptSegments.length > 0; }
    get chapterItems() {
        const result=[]; const seen=new Set();
        this.transcriptSegments.forEach((segment) => { const bucket=Math.floor(segment.startSeconds/300); if(!seen.has(bucket)){seen.add(bucket); result.push({key:`chapter-${bucket}`,time:segment.time,title:bucket===0?'Introduction':`Chapter ${bucket+1}`});} });
        return result;
    }
    get hasChapters() { return this.chapterItems.length > 0; }
    get acceptedCount() { return (this.details?.participants || []).filter((person) => person.responseToken === 'accepted').length; }
    get declinedCount() { return (this.details?.participants || []).filter((person) => person.responseToken === 'declined').length; }
    get attendanceRate() { return this.participantCount ? Math.round((this.attendedCount / this.participantCount) * 100) : 0; }
    get participantFilterOptions() { return [{label:'All participants',value:'all'},{label:'Attended',value:'attended'},{label:'Did not attend',value:'absent'},{label:'Accepted',value:'accepted'},{label:'Declined',value:'declined'},{label:'No response',value:'none'}]; }
    get visibleParticipants() {
        return (this.details?.participants || []).filter((person) => {
            const matchesText = !this.participantSearch || `${person.displayName} ${person.email} ${person.role}`.toLowerCase().includes(this.participantSearch);
            const matchesFilter = this.participantFilter === 'all' ||
                (this.participantFilter === 'attended' && person.actuallyJoined) ||
                (this.participantFilter === 'absent' && !person.actuallyJoined) ||
                person.responseToken === this.participantFilter;
            return matchesText && matchesFilter;
        });
    }
    get hasVisibleParticipants() { return this.visibleParticipants.length > 0; }

    mapDetails(result) {
        const meeting = result?.meeting || {};
        const start = this.parseLocal(meeting.startDateTime);
        const end = this.parseLocal(meeting.endDateTime);
        const emailCount = (meeting.attendeeEmails || '').split(/[;,]/).map((v) => v.trim()).filter(Boolean).length;
        const rawParticipants = result?.participants || [];
        const scheduledSeconds = start && end ? Math.max(0, Math.round((end - start) / 1000)) : 0;
        const participantEmails = new Set(rawParticipants.map((person) => (person.email || '').toLowerCase()));
        const invitedEmails = (meeting.attendeeEmails || '').split(/[;,]/).map((value) => value.trim()).filter(Boolean);
        const allParticipants = [...rawParticipants, ...invitedEmails.filter((email) => !participantEmails.has(email.toLowerCase())).map((email) => ({ email, displayName: this.nameFromEmail(email), invitationResponse: 'No response', actuallyJoined: false, intervals: [] }))];
        const participants = allParticipants.map((person, index) => this.mapParticipant(person, index, scheduledSeconds));
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
            relatedRecords: (result?.relatedRecords || []).map((record) => ({...record, key:`${record.objectApiName}-${record.recordId}`, label:record.recordLabel || record.recordId, objectLabel:(record.objectApiName || 'Record').replace(/__c$/,'').replaceAll('_',' ')})),
            participants
        };
    }
    close() { this.dispatchEvent(new CustomEvent('close')); }
    join() { if (this.meeting.joinUrl) window.open(this.meeting.joinUrl,'_blank','noopener'); }
    openOutlook() { if (this.meeting.webLink) window.open(this.meeting.webLink,'_blank','noopener'); }
    chooseTab(event) {
        const action = event.currentTarget.dataset.tab;
        if (['overview','participants','transcript','recording'].includes(action)) {
            this.activeTab = action;
            if (action === 'transcript' && !this.transcriptLoaded && this.transcriptAvailable) this.loadTranscript();
            return;
        }
        this.dispatchEvent(new CustomEvent('navigate', { detail: { action, meetingId: this.meetingId } }));
    }
    handleParticipantSearch(event) { this.participantSearch = (event.target.value || '').trim().toLowerCase(); }
    handleParticipantFilter(event) { this.participantFilter = event.detail.value; }
    handleTranscriptSearch(event) { this.transcriptSearch = (event.target.value || '').trim().toLowerCase(); }
    async loadTranscript() {
        this.transcriptLoading = true; this.transcriptError = undefined;
        try {
            const content = await getTranscriptContent({recordId:this.recordId,scheduleId:this.meetingId});
            this.transcriptSegments = this.parseWebVtt(content);
            this.transcriptLoaded = true;
        } catch(error) { this.transcriptError = this.reduceError(error); }
        finally { this.transcriptLoading = false; }
    }
    parseWebVtt(content) {
        const blocks=(content || '').replace(/\r/g,'').split(/\n\n+/); const rows=[];
        blocks.forEach((block,index) => { const lines=block.split('\n').filter(Boolean); const timingIndex=lines.findIndex((line)=>line.includes('-->')); if(timingIndex<0)return; const timing=lines[timingIndex].split('-->')[0].trim(); const text=lines.slice(timingIndex+1).join(' ').replace(/<[^>]+>/g,'').trim(); if(!text)return; const speakerMatch=text.match(/^([^:]{1,80}):\s*(.*)$/); const startSeconds=this.vttSeconds(timing); rows.push({key:`segment-${index}`,time:this.formatVttTime(startSeconds),startSeconds,speaker:speakerMatch?.[1] || 'Speaker',text:speakerMatch?.[2] || text,initials:this.initials(speakerMatch?.[1] || 'Speaker')}); });
        return rows;
    }
    vttSeconds(value){const parts=value.replace(',','.').split(':').map(Number); return parts.length===3?parts[0]*3600+parts[1]*60+parts[2]:parts[0]*60+parts[1];}
    formatVttTime(seconds){const minutes=Math.floor(seconds/60);return `${String(minutes).padStart(2,'0')}:${String(Math.floor(seconds%60)).padStart(2,'0')}`;}
    downloadTranscript() {
        const text=this.transcriptSegments.map((row)=>`${row.time} ${row.speaker}: ${row.text}`).join('\n\n');
        const url=URL.createObjectURL(new Blob([text],{type:'text/plain'})); const anchor=document.createElement('a'); anchor.href=url; anchor.download=`${this.meeting.subject || 'meeting'}-transcript.txt`; anchor.click(); URL.revokeObjectURL(url);
    }
    tabClass(name) { return this.activeTab === name ? 'active' : ''; }
    mapParticipant(person, index, scheduledSeconds) {
        const response = (person.invitationResponse || 'No response').toLowerCase();
        const responseToken = response.includes('accept') ? 'accepted' : response.includes('declin') ? 'declined' : response.includes('tentative') ? 'tentative' : 'none';
        const attendanceSeconds = Number(person.totalAttendanceSeconds || 0);
        const attendancePercent = person.actuallyJoined && scheduledSeconds ? Math.min(100, Math.round((attendanceSeconds / scheduledSeconds) * 100)) : 0;
        const intervals = person.intervals || [];
        return {...person, key:person.id || person.email || `participant-${index}`, displayName:person.displayName || this.nameFromEmail(person.email), initials:this.initials(person.displayName || person.email), role:person.role || person.attendeeType || 'Attendee', responseToken, responseLabel:responseToken === 'none' ? 'No response' : responseToken[0].toUpperCase()+responseToken.slice(1), responseClass:`pill response-${responseToken}`, attendanceLabel:person.actuallyJoined ? 'Attended' : 'Did not attend', attendanceClass:`pill ${person.actuallyJoined ? 'attendance-yes' : 'attendance-no'}`, joinedLabel:this.dateTime(person.firstJoinedAt || intervals[0]?.joinedAt), leftLabel:this.dateTime(person.lastLeftAt || intervals[intervals.length-1]?.leftAt), durationText:attendanceSeconds ? this.duration(attendanceSeconds) : '—', attendancePercent, progressStyle:`width:${attendancePercent}%`};
    }
    initials(value) { return (value || '?').split(/[@ ._-]/).filter(Boolean).slice(0,2).map((part) => part[0].toUpperCase()).join(''); }
    nameFromEmail(email) { return (email || 'Unknown attendee').split('@')[0].split(/[._-]/).filter(Boolean).map((part) => part[0].toUpperCase()+part.slice(1)).join(' '); }
    dateTime(value) { return value ? new Intl.DateTimeFormat('en-US',{hour:'numeric',minute:'2-digit'}).format(new Date(value)) : '—'; }
    duration(seconds) { const minutes=Math.floor(seconds/60); const secs=Math.round(seconds%60); return minutes ? `${minutes}m ${secs}s` : `${secs}s`; }
    moreAction(event) { this.dispatchEvent(new CustomEvent('action', { detail: { action:event.detail.value, meetingId:this.meetingId } })); }
    parseLocal(value) {
        const match = value?.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
        return match ? new Date(+match[1],+match[2]-1,+match[3],+match[4],+match[5],+(match[6]||0)) : null;
    }
    time(date) { return new Intl.DateTimeFormat('en-US',{hour:'numeric',minute:'2-digit'}).format(date); }
    plainText(html) { return (html || '').replace(/<br\s*\/?>/gi,' ').replace(/<[^>]+>/g,'').replace(/&nbsp;/g,' ').trim(); }
    reduceError(error) { return error?.body?.message || error?.message || 'The meeting details could not be loaded.'; }
}
