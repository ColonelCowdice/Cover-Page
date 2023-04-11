#Imports
import aspose.words as aw
import openai
import json

#API Imports
API_KEY = "Insert API Key Here"
openai.api_key = API_KEY

Name = input("What's your name?: ")
Age = input("What's your age?: ")
Company = input("Where do you want to work loser?: ")
Position = input("What position are you applying for?: ")
YearsExperience = input("Enter your years of experience in a similar career?: ")
Skill = input("What skills do you have that would make you a good candidate?: ")
EmailContact = input("What is your email: ")
PhoneContact = input("What is your phone number: ")



completion = openai.ChatCompletion.create(
  model="gpt-3.5-turbo",
  messages=[{"role": "user", "content": "Complete a cover letter for " + Name + " who is " + Age + " years old, using this available information, " + Company + ", " + Position + ", " + YearsExperience + ", " + Skill + ", " + EmailContact + ", " + PhoneContact + "."}],
)

completion = str(completion)
data = json.loads(completion)
content = data['choices'][0]['message']['content']

# create document object
doc = aw.Document()

# create a document builder object
builder = aw.DocumentBuilder(doc)

# create font
font = builder.font
font.size = 16
font.bold = True
font.name = "Arial"
font.underline = aw.Underline.DASH

# set paragraph formatting
paragraphFormat = builder.paragraph_format
paragraphFormat.first_line_indent = 8
paragraphFormat.alignment = aw.ParagraphAlignment.JUSTIFY
paragraphFormat.keep_together = True

# add text
builder.writeln(content)

# save document
doc.save("Tadah!.docx")
