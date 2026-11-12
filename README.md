# CanvQuizConverter

Write your quiz in Word. Get a file Canvas can import. That's the whole pitch.

**Use it:** https://techconsigliere.github.io/CanvQuizConverter/

Drop in a `.docx` quiz written in a simple numbered format, check the preview of every question and correct answer, and download a QTI `.zip` ready to import into a Canvas course. It runs entirely in your browser. The file is never uploaded, there's no account, and the page's Content Security Policy blocks it from making network requests to anyone. No Canvas admin access, no API keys.

The whole converter is one readable file, [`canvquiz.js`](canvquiz.js), using only what's built into the browser: no libraries.

## Who this is for

Anyone who teaches a Canvas course and would rather write a quiz the way they already write everything else: in Word, with a keyboard, without fighting a web form one question at a time.

Canvas admins are welcome to point faculty at it or fold it into onboarding materials. See the license note at the bottom.

## How to format your quiz

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

- Number each question: `1. Question text`. A question can run onto a second line.
- Letter each answer choice: `a. Choice text`. Upper or lowercase both work, and `1)` / `a)` work too. There's no limit at d.
- Put an asterisk before the letter of the correct answer: `*b. Choice text`.
- Mark more than one asterisk and the question becomes multiple-answer. The preview flags it so you can confirm.
- If Word turned your numbering into an automatic list, that's fine; the converter reads Word's numbering. Put the asterisk at the start of the choice text instead: `a. *Choice text`.

The preview flags a question with no correct answer, a question with fewer than two choices, and any line it couldn't place, rather than dropping it silently.

## Try the sample first

Download [`sample_quiz.docx`](sample_quiz.docx) and run it through before you write your own. It's a properly formatted eight-question quiz, including a true/false question. Fair warning: it's also a quiz about me, so you'll learn more than you expected to about your friendly neighborhood Canvas admin.

## Importing into Canvas

1. In your course, open **Settings**, then **Import Course Content** on the right.
2. Choose **QTI .zip file**, select the file you downloaded, and import.
3. The quiz appears under **Quizzes**, unpublished, with its questions also in a question bank. Preview it and set points, time limit, and dates before you publish.

Each download gets its own fresh internal ids, so importing a second quiz into the same course never lands on top of the first.

## What this doesn't do

It isn't a Canvas API integration. It doesn't connect to your course and can't push the quiz into Canvas for you. It produces a standard QTI package, the same import format Canvas already supports natively, and the import happens in Canvas itself. That's deliberate: no credentials change hands, and the tool can't touch anything beyond the file you give it.

It makes multiple choice, true/false, and multiple-answer questions for Classic Quizzes. Essay, matching, and fill-in-the-blank questions, images, and formatting in question text are left to Canvas.

## Run it on your own computer (Python)

The original version is a small Python program that runs the same conversion locally in a Streamlit window. Most people should use the web page above; this is here for anyone who wants it fully offline or wants to modify it. One known limit: the Python version reads only numbers and letters you typed yourself, not Word's automatic list numbering.

### Setup

You only need to do this once. After that, running the tool is a single command.

#### macOS

**1. Open Terminal.** Press `Command (⌘) + Spacebar`, type `Terminal`, press Enter.

**2. Install Python.** Recent versions of macOS no longer ship a usable Python, so go to [python.org/downloads](https://www.python.org/downloads/) and download the macOS installer. Run it, click through the installer (defaults are fine), and let it finish.

**3. Install the required libraries.** Copy and paste this into Terminal and press Enter:

```bash
pip3 install streamlit python-docx --break-system-packages
```

**4. Download CanvQuizConverter.** Go to the [CanvQuizConverter GitHub page](https://github.com/techconsigliere/CanvQuizConverter), click the green **Code** button, then **Download ZIP**. Unzip it, and you should have a file called `CanvQuizConverter.py`. Move it to your Desktop to make the next step easy.

**5. Run it.** In Terminal:

```bash
cd ~/Desktop
streamlit run CanvQuizConverter.py
```

Your browser will open automatically with the tool running.

#### Windows 11

**1. Install Python (first time only).** Go to [python.org/downloads](https://www.python.org/downloads/), click **Download Python**, and run the installer. On the very first screen, check the box that says **"Add python.exe to PATH"** before clicking Install Now. If you skip this box, none of the next steps will work.

**2. Open PowerShell.** Click Start, type `PowerShell`, right-click **Windows PowerShell**, and choose **Run as Administrator**.

**3. Install the required libraries.** Copy and paste this into PowerShell and press Enter:

```powershell
pip install streamlit python-docx
```

**4. Download CanvQuizConverter.** Go to the [CanvQuizConverter GitHub page](https://github.com/techconsigliere/CanvQuizConverter), click the green **Code** button, then **Download ZIP**. Unzip it, and you should have a file called `CanvQuizConverter.py`. Move it to your Desktop to make the next step easy.

**5. Run it.** In PowerShell:

```powershell
cd ~\Desktop
streamlit run CanvQuizConverter.py
```

Your browser will open automatically with the tool running.

#### Every time after that

You don't need to reinstall anything. Just open Terminal (Mac) or PowerShell (Windows), `cd` to wherever `CanvQuizConverter.py` lives, and run `streamlit run CanvQuizConverter.py` again.

To shut it down, click back into that black window and press `Control + C`.

### Troubleshooting

**macOS: "bad interpreter: ... no such file or directory"**

If `streamlit run CanvQuizConverter.py` throws an error like this, your Mac has a broken shortcut, not a broken tool. It usually shows up on managed campus machines after a Python or macOS update moves things around underneath an existing streamlit install. Run this instead:

```bash
python3 -m streamlit run CanvQuizConverter.py
```

That tells your current Python to load streamlit directly rather than trusting the old shortcut. It'll work every time, standard user rights or not, and there's nothing to fix or reinstall.

## Build

The page is built from readable source, never hand-edited inside its marked regions:

```
python3 build.py           # stamps VERSION into index.html and canvquiz.js, then inlines design/tc.css
python3 build.py --check   # exit 1 if anything is out of date
```

## License

CanvQuizConverter is released under the [MIT License](LICENSE). Use it, modify it, redistribute it, fold it into your campus onboarding materials, whatever you need. Releases before November 2026 were published under CC BY 4.0. Fira Sans is under the SIL Open Font License (`fonts/OFL.txt`).

Built by [Chris Powell](https://www.canvasinsider.blog), Canvas LMS Administrator, from the admin side of the desk.
