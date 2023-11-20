# google-script-calendar
An app script to add calendar events to Google calendar upon form submit.

This script has a form which upon being submitted triggers this script and adds a calendar event.
This script runs through the existing events looking for conflicts and adds the event to the calendar
once no conflicts are found.

It allows, one-shot events, daily recurring, weekly recurring, and monthy events. The form decides how many 
occurances can be added and that is something that can easily be increased or decreased by just editing the questions
in the form.

This script allows only adding and deleting as Google forms are not interactive, and it is very messy to create a workflow
that allow editing an event. This script also keeps it simple by asking the requester to store the event id generated post
addition to the calendar for use while deleting the event, if desired.
