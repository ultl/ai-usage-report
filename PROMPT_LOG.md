# Prompt Log — AI Usage Report Project

All user prompts from the development conversation, in chronological order.

---

## 2026-04-20 — Session 1

### Prompt 1
>
> Read the code and fix using openai compatible instead of ollama

### Prompt 2
>
> try to read data/template.xlsx and the code report.py, try to understand and modify the code. I want to show the est (without ai), actual (with ai) and time saved. Also I want to show the work of each person, for each task, and for each model they choose. At the end of the day, what I expected more is numbers. Ask me everything before implementing the code

### Prompt 3 (answers to clarification questions)
>
> 1. trust (user values), 2. All (all summary views), 3. Yes (keep AI inference), 4. Yes (want KPIs), 5. Keep as combined tool, 6. Both (xlsx + terminal)

### Prompt 4
>
> why when using python report.py data/*.xlsx -o report_16:17.xlsx --model gpt-5.4-mini, logging: Using OpenAI model: None

### Prompt 5
>
> I still use openai compatible

### Prompt 6
>
> why still 400?

### Prompt 7
>
> I want to create another python script to plot visual charts. Input data is the output of report.py. The charts plotted need to show: 1. So sánh số giờ est và actual khi dùng AI xem hiệu quả 2. Sự hài lòng khi dùng công cụ (đo bằng số lần rate *) 3. Dựa vào bài học người dùng và bài học AI suy luận, dùng openai compatible model để đọc và name các lỗi, và plot ra 1 chart so sánh những lỗi mà người dùng gặp phải... Cần follow best practice when prompt... Số lượng các lỗi cần có giới hạn và model cần phân loại chính xác xem mỗi người dùng gặp phải lỗi gì. Ví dụ: lỗi Clear and Format. Ask me everything before implement the code. I want the output to have a pretty format so that I can show to my CEO

### Prompt 8 (answer to design questions)
>
> You decide as I just want the format to be neat, clean and professional

### Prompt 9
>
> I want one more chart that show all tasks following up SDLC Stage

---

## 2026-04-20 — Session 2 (continued after context compaction)

### Prompt 10
>
> I want all files with .py to compact into only one python script and the outputs are what I received at all the current scripts

### Prompt 11
>
> can i remove all, except ai_journal.py

### Prompt 12
>
> at the current sheets, I see a lot of tables and charts that are duplicated, try dedup for me, you remove the sheet if in need

### Prompt 13
>
> I want to use LLM (model I provide in the command python ai_journal.py data/*.xlsx -o report.xlsx --model gpt-5.4-mini) to translate everything from input data to English. Output need to be all English

### Prompt 14
>
> Are you sure, the output I saw stil Vietnamese

### Prompt 15
>
> Check everything, I see report.xlxs still have vietnamses

### Prompt 16
>
> I want 3 sheets: per staff, per tool, per category into one sheet

### Prompt 17
>
> I want 3 sheets: rating charts, overview, breakdown, timetrend into one sheet

### Prompt 18
>
> In the error charts sheet, give description of each error label

### Prompt 19
>
> at input data, at "Nhật ký" sheet, there are EST (without AI) and Actual (with AI) which are provided by user, but I also want to use AI (I put in the commandline: python ai_journal.py data/sprint_0/*.xlsx -o report.xlsx --model gpt-5.4-mini) to give AI's perspective on the hours the work should be done without AI, and what should be done with AI. Of course, I need to provide the user's personnal for AI to understand user's background. How can I provide to you since each input file is each individual user?

### Prompt 20
>
> profiles.json

### Prompt 21
>
> I also want the comparison of user estimation and ai estimation at raw log, ai lesson compare sheets and at the pdf file

### Prompt 22
>
> I mean, user's estimation is a bit subjective. That's why I want to use AI to predict the AI EST, AI Actual, AI saved and I want to compare with user's estimation

### Prompt 23
>
> I want output presented in sheets and pdf file to give comparison between user's estimation and ai's estimation of working. Make sure that the report outputs need to show no subjectivity from user only, I want to use AI to make all this to be objective.

### Prompt 24
>
> efficiency at SDLC Summary need to be compared with AI estimation

### Prompt 25
>
> at SDLC summary sheet, I only want percentage of user estimation and ai estimation based on the tasks in the stage

### Prompt 26
>
> remove All Tasks Within Each SDLC Stage content and make the ai_journal_charts.pdf to have comparison between ai estimation and user estimation at SDLC stage to tasks + Efficiency

### Prompt 27
>
> Give short descriptions for all calculations

### Prompt 28
>
> from the output is report.xlsx and ai_journal_charts.pdf, create a detailed report of text with attached figures to show the highlights and significant points in the reports. Try to find out and ask me whether put the point in the report

### Prompt 29 (answer: which findings to include)
>
> all

### Prompt 30
>
> I mean to write a python script with using LLM to do task

### Prompt 31
>
> I want the report to explain the accuracy between what and what, like the user thinks the task is easier than it is

### Prompt 32
>
> change profile.json into profile.yml and modify the logic

### Prompt 33
>
> Create a markdown file logging all of my prompts with specific date in this conversation
