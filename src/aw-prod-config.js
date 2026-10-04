const testing = false
module.exports = {
    testing,
    // How far back to scrape call history (days) and an absolute safety cap on the
    // number of calls processed in a single run. See README for details.
    maxDaysBack: 30,
    maxCalls: 200,
    // Which calendar meetings to record. A meeting is kept only if its invitation
    // response is in meetingResponses AND its show-as (free/busy) is in meetingShowAs.
    // ResponseType: Accept, Organizer, Tentative, NoResponseReceived, Decline, Unknown
    // FreeBusyType: Busy, Tentative, Free, OOF (out-of-office), WorkingElsewhere
    meetingResponses: ['Accept', 'Organizer'],
    meetingShowAs: ['Busy']
}
