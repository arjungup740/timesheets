As a consultant I am required to fill out timesheets detailing the amount of time I've worked. Very rarely does this ever deviate from 40 hours. Thus I often end up (forgetting) to manually download a template sheet, update it for the current week, and email it.

Here I've written a simple python script that takes a template file, updates the relevants fields assuming 40 hours of work, then emails to me for review, so in 95% of cases I can simply forward on, instead of having to remember to find the template file and modify it.

I then containerize the application and put it on aws lambda. The lambda runs every morning via eventbridge. You can see this substack post here for the more in-depth discussion and code + AWS walkthrough!
https://quantitativecuriosity.substack.com/p/auto-fill-out-timesheets
