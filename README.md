# CanvQuizConverter

Write your quiz in Word. Get a file Canvas can import. That's the whole pitch.

CanvQuizConverter is a free tool that takes a quiz written in a Word document, using a simple numbered format, and converts it into a Canvas QTI package. Upload your `.docx`, preview your questions and correct answers in the browser, and download a `.zip` file ready to import directly into a Canvas course.

No Canvas admin access required. No API keys. No account. It runs entirely on your own computer, and nothing you upload ever leaves your machine.

## Who this is for

Anyone who teaches a Canvas course and would rather write a quiz the way they already write everything else: in Word, with a keyboard, without fighting a web form one question at a time.

Canvas admins are also welcome to point faculty at this, host it on a lab machine, or fold it into onboarding materials. See the license note at the bottom.

## How to format your quiz

Open a blank Word document and write your quiz like this:

```
1. Chris' favorite color is:
*a. Black
b. Blue
c. Red
d. Turquoise

8. Chris is happy to convert your properly-formatted Word document into a quiz or question bank in your Canvas course.
*a. True
b. False
```

A few rules that matter:

- Number each question, period, then the question text: `1. Question text`
- Each answer choice gets a letter, period, then the choice text: `a. Choice text`
- Put an asterisk directly before the letter of the correct answer: `*b. Choice text`
- Mark more than one asterisk on a question and the tool will treat it as a multiple answer question automatically. It'll flag this in the preview so you can confirm that's what you meant.
- Letters can be upper or lowercase. The tool doesn't care.

That's it. No styles, no tables, no special formatting required.

## Not sure if you've got it right? Try the sample first.

Download [`sample_quiz.docx`](https://github.com/techconsigliere/CanvQuizConverter/blob/main/sample_quiz.docx) and run it through the tool before you write your own. It's a properly-formatted eight-question quiz, including a true/false question, so you can see exactly what correct formatting looks like and confirm the tool is working before you invest time writing your real one. Fair warning: it's also a quiz about me, so you'll learn more than you expected to about your friendly neighborhood Canvas admin in the process.

## Installation and setup

You only need to do this once. After that, running the tool is a single command.

### macOS

**1. Open Terminal.** Press `Command (⌘) + Spacebar`, type `Terminal`, press Enter.

**2. Install Python.** Recent versions of macOS no longer ship a usable Python, so go to [python.org/downloads](https://www.python.org/downloads/) and download the macOS installer. Run it, click through the installer (defaults are fine), and let it finish.

**3. Install the required libraries.** Copy and paste this into Terminal and press Enter:

```bash
pip3 install streamlit python-docx --break-system-packages
```

**4. Download CanvQuizConverter.** Go to the [CanvQuizConverter GitHub page](https://github.com/techconsigliere/CanvQuizConverter), click the green **Code** button, then **Download ZIP**. Unzip it, and you should have a file called `quizconverter.py`. Move it to your Desktop to make the next step easy.

**5. Run it.** In Terminal:

```bash
cd ~/Desktop
streamlit run quizconverter.py
```

Your browser will open automatically with the tool running.

### Windows 11

**1. Install Python (first time only).** Go to [python.org/downloads](https://www.python.org/downloads/), click **Download Python**, and run the installer. On the very first screen, check the box that says **"Add python.exe to PATH"** before clicking Install Now. If you skip this box, none of the next steps will work.

**2. Open PowerShell.** Click Start, type `PowerShell`, right-click **Windows PowerShell**, and choose **Run as Administrator**.

**3. Install the required libraries.** Copy and paste this into PowerShell and press Enter:

```powershell
pip install streamlit python-docx
```

**4. Download CanvQuizConverter.** Go to the [CanvQuizConverter GitHub page](https://github.com/techconsigliere/CanvQuizConverter), click the green **Code** button, then **Download ZIP**. Unzip it, and you should have a file called `quizconverter.py`. Move it to your Desktop to make the next step easy.

**5. Run it.** In PowerShell:

```powershell
cd ~\Desktop
streamlit run quizconverter.py
```

Your browser will open automatically with the tool running.

### Every time after that

You don't need to reinstall anything. Just open Terminal (Mac) or PowerShell (Windows), `cd` to wherever `quizconverter.py` lives, and run `streamlit run quizconverter.py` again.

To shut it down, click back into that black window and press `Control + C`.

## Troubleshooting

**macOS: "bad interpreter: ... no such file or directory"**

If `streamlit run quizconverter.py` throws an error like this, your Mac has a broken shortcut, not a broken tool. It usually shows up on managed campus machines after a Python or macOS update moves things around underneath an existing streamlit install. Run this instead:

```bash
python3 -m streamlit run quizconverter.py
```

That tells your current Python to load streamlit directly rather than trusting the old shortcut. It'll work every time, standard user rights or not, and there's nothing to fix or reinstall.

## Using the tool

1. Upload your `.docx` quiz file.
2. Review the preview. Correct answers show in green. Anything missing or ambiguous gets flagged.
3. Name your export file.
4. Click **Convert & Download**.
5. In Canvas, go to your course, click **Settings** in the left navigation, then the **Import Course Content** link on the right. Choose **QTI .zip file**, select the file you just downloaded, and import.

Your quiz will appear in the course's question bank and as a new quiz, ready to assign.

## A note on what this doesn't do

This isn't a Canvas API integration. It doesn't connect to your course, doesn't know your enrollment, and can't push the quiz into Canvas for you. It produces a standard QTI package, the same import format Canvas already supports natively, and the import step happens in Canvas itself. That's deliberate. No credentials change hands, and the tool has no way to touch anything beyond the file you feed it.

## License

CanvQuizConverter is released under [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). Use it, modify it, redistribute it, fold it into your institution's onboarding materials, whatever you need. All that's asked is attribution back to the original source.

Built by [Chris Powell](https://canvasinsider.blog), Canvas LMS Administrator, from the admin side of the desk.
