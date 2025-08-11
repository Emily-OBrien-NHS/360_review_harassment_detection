##########TO DO:
# - SEE IF CAN SPEED UP RUN TIME
# - AUTOMATE RUNNING EVERY 2 WEEKS
# - ADD TO GITHUB

import re
import os
import time
import numpy as np
import pandas as pd
from sqlalchemy import create_engine
from detoxify import Detoxify
import win32com.client as win32
from datetime import datetime
#https://pypi.org/project/detoxify/
from better_profanity import profanity
#https://pypi.org/project/better-profanity/
import nltk
from nltk.stem import WordNetLemmatizer
nltk.download("wordnet")
nltk.download("omw-1.4")

################################################################################
                        #Read and Pre-Process Data#
################################################################################
#Initialising parameters and models
t0 = time.time()
today = datetime.today().strftime('%Y-%m-%d')
wnl = WordNetLemmatizer()
original = Detoxify('original')
os.chdir('C:/Users/obriene/Projects/Text Data Analysis/360 Feedback Analysis')
print('Reading in Data')

#Read in 360 feedback responses
AG_DW_PROD_engine = create_engine('mssql+pyodbc://@AG-DW-PROD/RPA?'\
                           'trusted_connection=yes&driver=ODBC+Driver+17'\
                               '+for+SQL+Server')
feedback_sql = '''SELECT ID, CompletionTime, EmployeeNumber, PersonLeader,
                         TrustStrategy, StrategicPlanning, Communicates,
                         TwoWayCommunication, RespondFeedback, PositiveChange
                  FROM [RPA].[UHPForm].[UHP106]'''
feedback = pd.read_sql(feedback_sql, AG_DW_PROD_engine)
print('Processing Data')

#Filter out completed IDs
complete_IDs_excel = pd.read_excel('Completed IDs.xlsx')
complete_IDs = complete_IDs_excel['Complete IDs'].to_list()
feedback = feedback.loc[~feedback['ID'].isin(complete_IDs)].copy()
print(f'There are {len(feedback)} new responses to process')

#List of the columns with text responses
text_cols = ['TrustStrategy', 'StrategicPlanning', 'Communicates',
             'TwoWayCommunication', 'RespondFeedback', 'PositiveChange']

#Get list of offensive, agressive and hateful words from
#https://github.com/dsojevic/profanity-list/blob/main/en.json, put them
#together into one big regex string to search for
profanitys = pd.read_json('en.json')
#explode out all the matches to one row and filter out * (due to regex complications)
profanitys['match'] = profanitys['match'].str.split('|')
profanitys = profanitys.explode('match')
profanitys = profanitys.loc[~profanitys['match'].str.contains(r'\*', regex=True)].copy()
profanitys['tags'] = [', '.join(tag) for tag in profanitys['tags']]

#Add on list of words from caroline
uhp_list = pd.read_excel('20250729 Terms for Sentiment Analysis.xlsx')
uhp_list['tags'] = uhp_list['tags'].ffill()
uhp_list = uhp_list.dropna()
#join together and create regex
word_list = pd.concat([profanitys[['id', 'match', 'tags']], uhp_list])
word_list = word_list.drop_duplicates(subset='id', keep='first')
word_list['match'] = word_list['match'].fillna(word_list['id'])
regex = r'\b' + r'|'.join(word_list['match']).replace('|', r'\b|\b') + r'\b'

################################################################################
                        #Look for offensive responses#
################################################################################
print('Looking for offensive responses')
#Go through each text column and evaluate
for col in text_cols:
    texts = feedback[col]
    #Emptyy lists to store found abusive words, if profanity is found and what
    #tags these fall under
    ab_words = []
    contains_profanity = []
    tags = []
    #Go thorugh each text and flag words and profanity
    for text in texts.values:
        #create a lemmatized version of the text and add to the end of the OG
        #text to ensure any conjugated words still match
        lemmatized_text = ' ' + ' '.join([wnl.lemmatize(word)
                                          for word in text.split(' ')])
        text += lemmatized_text
        text_tags = []
        #####Word flagging#####
        flagged_words = list(set([word for word in re.findall(regex, text.lower())
                         if word != 'sit']))
        ab_words.append(flagged_words)
        if flagged_words:
            #If words found, record what label they come under
            text_tags += word_list.loc[word_list['match']
                                       .str.contains(r'\b'
                                                     + r'\b|\b'.join(flagged_words)
                                                     + r'\b', regex=True),
                                       'tags'].tolist()
        #####Profanity Flagging#####
        profanity_flag = profanity.contains_profanity(text)
        contains_profanity.append(profanity_flag)
        if profanity_flag:
            #If profanity is found, record this
            text_tags.append('Profanity')
        #####Asterisk Flag#####
        if '*' in text:
            text_tags.append('Aterisk Check')
        tags += [text_tags]

    #####Using model####
    results = pd.DataFrame(original.predict(texts.values.tolist()))
    #Sum up the outputted offense probabilities, and apply the most
    #dominant label
    sums = results.sum(axis=1)
    #dataframe to store the 3 different flags for recording.
    df = pd.DataFrame({'detoxify' : [i > 0.5 for i in sums.to_list()],
                       'profanity' : contains_profanity,
                       'abusive' : [len(i) > 0 for i in ab_words]})
    #Add detoxify to the tags if this flags
    tags = [(tags[i].append('Detoxify') if df.loc[i, 'detoxify'] else tags[i])
            for i in range(len(tags))]
    tags = [(tag if tag else np.nan) for tag in tags]
    #Add flag and reason columns to the output
    feedback[col + ' FLAG'] = df.any(axis=1)
    feedback[col + ' FLAG reason'] = tags
    feedback[col + ' FLAG words'] = ab_words

################################################################################
                        #Create Final Table to Output#
################################################################################
#Create overall flag column if any of the responses flag in that row
feedback['FLAG'] = feedback[[col for col in feedback.columns
                             if 'FLAG' in col]].any(axis=1)
feedback['REASONS'] = {q: s.replace([], np.nan).dropna().to_dict()
                       for q, s in feedback[[col for col in feedback.columns
                                             if 'reason' in col]].iterrows()}
print(f'There are {feedback['FLAG'].sum()} flagged responses')

#Sort columns
feedback_cols = (['ID', 'CompletionTime', 'EmployeeNumber', 'PersonLeader']
                 + sorted([col for col in feedback.columns if
                           any(text_col in col for text_col in text_cols)])
                 + ['FLAG', 'REASONS'])
feedback = feedback[feedback_cols].reset_index(drop=True).copy()


################################################################################
                            #Save to Excel#
################################################################################
file_path = f'C:/Users/obriene/Projects/Text Data Analysis/360 Feedback Analysis/Outputs/Flagged 360 responses {today}.xlsx'
writer = pd.ExcelWriter(file_path, engine="xlsxwriter")
feedback[feedback_cols].to_excel(writer, sheet_name='Responses', index=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets["Responses"]

# Add formats
text_format = workbook.add_format({'align' : 'centre'})
text_wrap_format = workbook.add_format({'text_wrap' : True, 'align' : 'left'})

# Set columns.
#initial data
worksheet.set_column('A:A', 5, text_format)
worksheet.set_column('B:C', 20, text_format)
worksheet.set_column('D:D', 30, text_format)
#communicates
worksheet.set_column('E:E', 50, text_wrap_format)
worksheet.set_column('F:F', 15, text_format)
worksheet.set_column('G:H', 14, text_wrap_format)
#Positive Change
worksheet.set_column('I:I', 50, text_wrap_format)
worksheet.set_column('J:J', 15, text_format)
worksheet.set_column('K:L', 14, text_wrap_format)
#Respond to Feedback
worksheet.set_column('M:M', 50, text_wrap_format)
worksheet.set_column('N:N', 15, text_format)
worksheet.set_column('O:P', 14, text_wrap_format)
#Strategic Planning
worksheet.set_column('Q:Q', 50, text_wrap_format)
worksheet.set_column('R:R', 15, text_format)
worksheet.set_column('S:T', 14, text_wrap_format)
#Trust Strategy
worksheet.set_column('U:U', 50, text_wrap_format)
worksheet.set_column('V:V', 15, text_format)
worksheet.set_column('W:X', 14, text_wrap_format)
#Two-way Communication
worksheet.set_column('Y:Y', 50, text_wrap_format)
worksheet.set_column('Z:Z', 15, text_format)
worksheet.set_column('AA:AB', 14, text_wrap_format)
#FLAG
worksheet.set_column('AC:AC', 8, text_format)
worksheet.set_column('AD:AD', 50, text_wrap_format)

# Set the autofilter.
(max_row, max_col) = feedback.shape
worksheet.autofilter(0, 0, max_row, max_col)

# Filter to where Flags is true and hid the rows where false
worksheet.filter_column("AC", "X == TRUE")
for row_num in feedback.loc[~feedback["FLAG"]].index.tolist():
    worksheet.set_row(row_num + 1, options={"hidden": True})

# Close the Pandas Excel writer and output the Excel file.
writer.close()

################################################################################
                                #Send Email#
################################################################################
#send email with latest flagged output attatched    
# Create Outlook application object and mail item
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)    
# Set email properties
mail.To = open('emails.txt', 'r').read()
#mail.To = 'e.obrien6@nhs.net'
mail.Subject = '360-feedback flagged responses'
mail.Body = """Hi,\n
Please find attatched the latest flagged responses from the 360-degree feedback forms. Let me know if you have any issues or questions \n
Emily"""
#Attatch the flagged file
mail.Attachments.Add(file_path)
# Send email
mail.Send()
print(f"Email sent successfully")

################################################################################
                        #Update processed IDs list#
################################################################################
#Update the complete IDs list with IDs that have been processed
print('Updating completed IDs')
complete_IDs_excel = pd.concat([complete_IDs_excel,
                                pd.DataFrame(feedback['ID'].values,
                                             columns=['Complete IDs'])]
                              ).drop_duplicates()
complete_IDs_excel.to_excel('Completed IDs.xlsx', index=False)
t1 = time.time()
print(f'Complete in {((t1-t0)/60):.2f} mins')
