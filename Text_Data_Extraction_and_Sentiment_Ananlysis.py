# Importing all necessary libaryes
import copy
import string
import os
from playwright.sync_api import sync_playwright
import pandas as pd
import nltk
# nltk.download('punkt')
from nltk.tokenize import RegexpTokenizer, BlanklineTokenizer, sent_tokenize
from nltk.corpus import wordnet
import pandas as pd
# nltk.download('wordnet')
all_complex_words = set(wordnet.words())

# Essential functions for code

# You can have your own stop words or call stopwords of NLTK but in this code stopwords are given in text file
stopwords_folder_path = r"F:\Codes"  # Enter folder location containing stopword files
stopwords_files = [file for file in os.listdir(stopwords_folder_path)]
stopwords = []
for stopwords_file in stopwords_files:
    location = f"{stopwords_folder_path}\\{stopwords_file}"
    with open(location, 'r') as swf:
        sf = swf.read()
    
    # normalizing word to lower
    sf = sf.lower()
    tokenizer = RegexpTokenizer(r'\w+')
    sf_list = tokenizer.tokenize(sf)
    stopwords.extend(sf_list)

stopwords = set(stopwords)

# Cleaning of text using stopwords
def cleaning_stopwords(x):
    x_copy = copy.deepcopy(x)
    for i in stopwords:
        while i in x_copy:
            x_copy.remove(i)
    return(len(x_copy))

# Creating a dicnary of positive and negative words
path_positive_words = r"F:\Codes\positivewords.txt"  # Give file location containing positive words
path_negative_words = r"F:\Codes\negativewords.txt"  # Give file location containing negative words

with open(path_positive_words, 'r') as ppw:
    positive_words = ppw.read()
positive_words = nltk.word_tokenize(positive_words) # converting input text to list from string

#Dictnary of positive words
def positive_number_count(xw):
    positive_word_dict = {}
    for i in positive_words:
        positive_word_dict[i] = xw.count(i)
    return sum(positive_word_dict.values())


with open(path_negative_words, 'r') as ppw:
    negative_words = ppw.read()
negative_words = nltk.word_tokenize(negative_words)

# #Dictnary of negative words
def negative_number_count(xw):
    negative_word_dict = {}
    for i in negative_words:
        negative_word_dict[i] = xw.count(i)
    return sum(negative_word_dict.values())

# Columns for excel 'Input all columns(functions) you need to get from program'
temp_df = {'URL_ID' : [],
    'URL' : [],
    'POSITIVE SCORE' : [],  
    'NEGATIVE SCORE' : [], 
    'POLARITY SCORE' : [], 
    'SUBJECTIVITY SCORE' : [], 
    'AVG SENTENCE LENGTH' : [], 
    'PERCENTAGE OF COMPLEX WORDS' : [], 
    'AVG NUMBER OF WORDS PER SENTENCE' : [],
    'COMPLEX WORD COUNT' : [],
    'WORD COUNT' : [],
    'SYLLABLE PER WORD' : [],
    'PERSONAL PRONOUNS' : []
    }


# Extracting data from input excel file links

input_urls = pd.read_excel(r"F:\Codes\Input_urls.xlsx")   # Input the excel sheet containg urls for better understanding put name of each urls
# print(input_urls)
for index, row in input_urls.iterrows():
    # variable to store text
    t = ''
    i = row['URL_ID']
    i1 = str(i)
    index_deci = i1.index('.')
    if int(i1[index_deci+1:]) == 0:
        i = int(i)
    j = row['URL']  

    # print(i, type(i))
    
    url = j
    output_directory = r"F:\Codes\text_files"   # Here program will store text data extracted from url 
    file_name = f"{i}.txt"
    output_file_name = f"{output_directory}\\{file_name}"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  #Launching Chrome webdriver
        page = browser.new_page()                    # Opening a new page
        page.goto(url)
        #Increasing loading time
        page.set_default_timeout(10000)
        title = page.title()                         # Extracting Page Title

        # print(i, title)

        element = page.locator('.example_xyz')      # Input element that you want to extract    
        with open(output_file_name, 'w', encoding='utf-8') as output_file:
            output_file.write(f"{title}\n\n")

            for paragraph in element.locator('p').all():
                paragraph_text = paragraph.text_content()
                output_file.write(paragraph_text + "\n")
                t += paragraph_text

        browser.close()

    # Calculating all variables values
    # saving text file with casesensitive
    t1 = t.strip()
    # Normalizing words to lower format
    t = t.strip().lower()

    word_list = nltk.word_tokenize(t)
    word_list = [word for word in word_list if word not in string.punctuation]
    # print(len(word_list))

    # for page not found
    if len(word_list) ==0:
        positive_score = 0
        negative_score = 0
        polarity_score = 0
        subjectivity_score = 0
        avg_sent_len = 0
        percentage_complex_words = 0
        fog_index = 0
        avg_words_sentence = 0
        complex_word_count = 0
        cleaned_word_count = 0
        syllable_count = 0
        personal_pernaunce_count = 0
        avg_word_length = 0
    
    else:
        # 1.3.a Positive score
        positive_score = positive_number_count(word_list)
        # print(positive_score, 'p')

        #1.3.b Negative score
        negative_score = negative_number_count(word_list)
        # print(negative_score, 'n')

        #Polarity Score
        polarity_score = (positive_score - negative_score)/ ((positive_score + negative_score) + 0.000001)
        # print(polarity_score, 'ps')

        # subjectivity score
        total_words_after_clening = cleaning_stopwords(word_list)
        subjectivity_score = (positive_score + negative_score)/ ((total_words_after_clening) + 0.000001)

        #2.1 Average sentence length
        total_number_sentense = t.replace('\n',' ')
        total_number_sentense = nltk.sent_tokenize(total_number_sentense) 
        avg_sent_len = len(word_list)/len(total_number_sentense)
        # print(avg_sent_len,'asl')

        #2.2 percentage of complex words
        total_complex_words = [i for i in word_list if i in all_complex_words]
        percentage_complex_words = len(total_complex_words)/len(word_list)
        # print(percentage_complex_words, 'pcw')

        #2.3 Fog index
        fog_index = 0.4*(avg_sent_len + percentage_complex_words)
        # print(fog_index,'fi')

        #3 Average number of words per sentence
        avg_words_sentence = len(word_list)/len(total_number_sentense)
        # print(avg_words_sentence, 'aws')

        #4 Complex word count
        complex_word_count = len(total_complex_words)
        # print(complex_word_count, 'cws')

        #5 Word count without stop words and punctuation
        cleaned_word_count = cleaning_stopwords(word_list)
        # print(cleaned_word_count, 'c1ws')


        #6 syllabe count per word

        exception = ['es', 'ed']              # input exceptions for syllabe count
        vowels_counts = {}.fromkeys('aeiou', 0)

        for word in word_list:
            # Check if the word ends with any of the exceptions
            if not any(word.endswith(ex) for ex in exception):

                vowels_counts['a'] += word.count('a')
                vowels_counts['e'] += word.count('e')
                vowels_counts['i'] += word.count('i')
                vowels_counts['o'] += word.count('o')
                vowels_counts['u'] += word.count('u')

        syllable_count = sum(vowels_counts.values())
        # print(syllable_count, 'sc')


        #7 Personal pernaunce
        case_sensitive_word_list = nltk.word_tokenize(t1)
        case_sensitive_word_list = [case_sensitive_word for case_sensitive_word in case_sensitive_word_list if case_sensitive_word not in string.punctuation]
        # print(case_sensitive_word_list)
        case_word_dict = {}
        case_word_dict['I'] = case_sensitive_word_list.count('I')
        case_word_dict['we'] = case_sensitive_word_list.count('we')
        case_word_dict['my'] = case_sensitive_word_list.count('my')
        case_word_dict['ours'] = case_sensitive_word_list.count('ours')
        case_word_dict['us'] = case_sensitive_word_list.count('us')

        personal_pernaunce_count = sum(case_word_dict.values())
        # print(personal_pernaunce_count, 'ppc')

        #8 Average word length
        total_number_of_characters = "".join(word_list)
        avg_word_length = len(total_number_of_characters)/len(word_list)
        # print(avg_word_length, 'awl')

    # Adding values to dictnary
    temp_df['URL_ID'].append(i)
    temp_df['URL'].append(j)
    temp_df['POSITIVE SCORE'].append(positive_score)
    temp_df['NEGATIVE SCORE'].append(negative_score)
    temp_df['POLARITY SCORE'].append(polarity_score)
    temp_df['SUBJECTIVITY SCORE'].append(subjectivity_score)
    temp_df['AVG SENTENCE LENGTH'].append(avg_sent_len)
    temp_df['PERCENTAGE OF COMPLEX WORDS'].append(percentage_complex_words)
    temp_df['FOG INDEX'].append(fog_index)
    temp_df['AVG NUMBER OF WORDS PER SENTENCE'].append(avg_words_sentence)
    temp_df['COMPLEX WORD COUNT'].append(complex_word_count)
    temp_df['WORD COUNT'].append(cleaned_word_count)
    temp_df['SYLLABLE PER WORD'].append(syllable_count)
    temp_df['PERSONAL PRONOUNS'].append(personal_pernaunce_count)
    temp_df['AVG WORD LENGTH'].append(avg_word_length)

# saving dataframe to excel file
df = pd.DataFrame(temp_df)
df.to_excel("F:/Codes/Output Data.xlsx", index=False)      # Save data in Excel file
