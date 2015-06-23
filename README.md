# aspiring-ninja

<h1>Summary</h1>
<p>At AOL, we have been using a monthly self-assessment and review process with the Paid Services Engineering team since May 2015. Prior to that time we had been evaluating a third party a performance management tool among our leadership team. I noted that we were really only using the performance review features of the tool, so when my manager asked me my opinion on whether we should purchase this tool for use with our entire team I told him "No. I think I can build what we need." Hence this project was born.</p>

<p>The tool itself is built using Google Apps and Google Apps Script. When I started, I had a simple Google Form, a Google Sheet for capturing responses, a Google Doc-based Performance Review template, and a Google Apps Script to copy the template, populate it with individual form responses, rename it, save, it and share it with the employee and manager. Employees entered their names, selected their manager's name form a drop-down list, and all tracking was done manually.</p>

<p>Since that initial incarnation, I have added an automated submission tracker, an automated email reminder series, and eliminated the need for employees to provide their names or manager information by incorporating an employee roster that leverages our enterprise Google Apps environment to look that information up based on the employee's email address. I am providing my code here for anyone who may be interested in trying a similar approach.<p>

<h1>Getting Started</h1>

<p>There are a few documents and forms that I use to make the whole thing "go."</p>

<ul>
<li><a href = "https://docs.google.com/forms/d/14o3W566F9ING51U-IjL0dwF4P_lzTcQotIxfJKNQStg/viewform?usp=send_form">Monthly Self-Assessment Form</a></li>
<li><a href = "https://docs.google.com/forms/d/14o3W566F9ING51U-IjL0dwF4P_lzTcQotIxfJKNQStg/viewform?usp=send_form">Monthly Review Document (template)</a></li>
<li><a href = "https://docs.google.com/spreadsheets/d/1Icv6RPv_fvO_vissV8sMXtW7fpVGP9e4UKJccmgSHzE/edit?usp=sharing">Employee Roster (with management chain information)</a></li>
<li><a href = "https://docs.google.com/spreadsheets/d/1DNJlze_VdtSuHp4zSlIDhzMOGZ9CxB-_4SZzFtQcYLI/edit?usp=sharing">Assessment Submission Tracker</a></li>
</ul>

<p>If you'd like to replicate or modify my process, start by making copies of each of these documents. You can then grab the script I have provided and modify the code to refer to your documents.</p>

<h1>Process Overview</h1>

<p>Form submissions are captured in a spreadsheet that only the senior leader in our team's management chain has access to and there is a script embedded in that sheet which generates review documents based on the template, sends out submission confirmation emails, and updates the submission tracker. Embedded within the submission tracker is a reminder email series that I have automated to be delivered monthly (I can provide additional detail if desired) - check out the Reminder tab for more information</p>


<ol>
<li>Reminder email sent to all employees on first business day of each month to complete self-assessment (link to form included)</li>
<li>Employees have 3 business days to complete their self-assessment using the form above (this can be customized)</li>
<li>Reminder email sent to people managers that review cycle is underway (includes a link to the submission tracker)</li>
<li> When an employee submits a self-assessment two emails are triggered - one to the employee with a link to their self-assessment document (view only) and one to the employees direct manager and 2d level manager with a link to the self-assessment (I use a CC list as well so I am copied on each of the manager emails). This is purely optional.</li>
<li>A reminder email to schedule 1:1 conversations is sent to people managers the first business day after the self-assessment period has "ended."</li>
<li>Direct managers have edit rights to self-assessments submitted by their employees. They should add comments to each document to close the feedback loop - this process is currently self-regulated. Comments can be added at any time - either in advance of the 1:1 conversation or immediately after. 2d level managers have view access to self-assessments submitted by employees in their management chain.</li>
