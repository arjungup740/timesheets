As a consultant I am required to fill out timesheets detailing the amount of time I've worked, very rarely does this ever deviate from 40 hours. Thus I often end up (forgetting) to manually download a template sheet, update it for the current week, and email it.

Here I've written a simple python script that takes a template file, updates the relevants fields assuming 40 hours of work, then emails to me for review, so in 95% of cases I can simply forward on, instead of having to remember to find the template file and modify it.

I then containerize the application and put it on aws lambda. The lambda runs every morning via eventbridge

TODOs:
* ultimately could have some logic that checks which emails you sent previously
* add some boto that pulls your costs
* use google oauth instead of creating an app password
* have system email you and ask how if you worked 40 hours, if reply yes, then send 40 hour email directly. If no, then send a filled out timesheet saying "please modify"

done
1) figure out date logic -- handle if you send previous email or not.... or just default to being careful and only sending one for the previous week unless it's Friday late pm the previous week
    if it is friday afternoon, send for last week. If it is anytime before Friday afternoon, send for the previous week. If it is Monday or any other day really, send for 2 saturdays ago, because presumably you missed it
* containerize so it just always works
* run via lambda so that doesn't matter if your computer is on
* set up lambda timing so it fires weekly
* change the secrets
"""