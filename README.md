# CV-Screener
A CV Screener using OpenAI API that I solely developed and created at BrightIsle in January 2025

*This program depends on the file naming schema [Resume/CoverLetter]\_[First]-[Last]\_[Source].pdf because that's how they are downloaded at brightisle*

How to add your OpenAI API key:
1. On the first run, a file named config.json will be created in the directory: C:\Users\YOURNAME\AppData\Roaming\BrightIsle CV Screener
2. Open the file config.json using any text editor.
3. Enter your OpenAI api key as shown below:
{
    "OPENAI_API_KEY": "your-api-key-here"
}
4. Save the file and restart the application.

If you delete or edit the config.json file, the program will not work. To restore the file, delete it (if edited), and run the application again and proceed from step 1. 



How to use:
1. Double click the dropbox to add Resumes and Coverletters that you want processeed. Only writeable PDFs are valid file types, camscanned PDFs are NOT supported. Uploading anything else will result in an error.
2. Add keywords and guidelines into Screening Criteria. The AI used will be given a prompt which includes user criteria as key features to highlight. The more criteria a resume meets
the higher it will score. Since an AI is interpreting the resumes, the keywords given do not need to be directly included in the resumes, it will just use your instructions to interpret them.
3. Change AI strength scale. The strength will determine how many resumes you would like to manually review based on your criteria. Below is the prompt that we use to determine filter strength.


Filter Strength Strictness:
1 - Very low strength, Only disqualify applicants who are completely unqualified and do not meet any criteria.
2 - Low strength, Disqualify applicants who are underqualified and do not adequately meet most criteria.
3 - Medium strength, Disqualify applicants who meet only some criteria but are still slightly underqualified.
4 - High strength, Disqualify applicants who meet the minimum criteria but are not strong candidates overall.
5 - Very high strength, Only approve applicants who perfectly or nearly perfectly match the criteria and are standout candidates.
