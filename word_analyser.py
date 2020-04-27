# -*- coding: utf-8 -*-
"""
Created on Fri Apr 17 19:20:28 2020

@author: 
"""
import pandas as pd
import matplotlib.pyplot as plt
import math
# import nltk
#nltk.download()
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize

#read excel
# convert the dictionary to list and remove the unicode from string
def read_excel_and_format(excel_file, sheet_number):
    """
    

    Parameters
    ----------
    excel_file : excel file (xlsx)
        name of the excel file (.xlsx
    sheet_number : int
        sheet number of the excel file to read
    Returns
    -------
    list
        data from the excel file.

    """
    #use pandas to read the excel and oobtain a dataframe
    df = pd.read_excel(excel_file,sheet_number)
    #convert the dataframe to list
    lst = df.values.tolist()
    #variable that holds the formated lists
    formated_lst = []
    #loop through the data and format the data then append to a list
    for ls in lst:
        #temporary sublist
        sub_lst = []
        #change the encoding
        str_org = ls[2].encode('ascii','ignore')
        str_functionality = ls[3].encode('ascii','ignore')
        str_owner = ls[4].encode('ascii','ignore')
        #append to the temp sublist
        sub_lst.append(ls[0])
        sub_lst.append(ls[1])
        sub_lst.append(str_org)
        sub_lst.append(str_functionality)
        sub_lst.append(str_owner)

        formated_lst.append(sub_lst)
    return formated_lst


# remove comment and convert to format words in functionality
def formart_functionality(formated_lst):
    """   

    Parameters
    ----------
    formated_lst : List
        List of word read from excel file.

    Returns
    -------
    no_comments : list
        Cleaned list without comment section.

    """
    no_comments = []
    #loop through the list as you format the data
    for item in formated_lst:
        #remove the comments
        lines_of_words_minus_comments = item[3].split('Comment')[0]
        #remove punctuations
        a = lines_of_words_minus_comments.replace(".", '')
        b = a.replace(";", '')
        c = b.replace(",", '')
        d = c.replace("\"", '')
        converted_words_to_lower = d.lower()

        sub_lst = []
        sub_lst.append(item[0])
        sub_lst.append(item[1])
        sub_lst.append(item[2])
        sub_lst.append(converted_words_to_lower)
        sub_lst.append(item[4])

        no_comments.append(sub_lst)
    return no_comments
    

#count the word frequency, output format
# word: [freqency, organization, user]
def count_words(no_comments):
    """
    Parameters
    ----------
    no_comments : list
        words that have been formatted.

    Returns
    -------
    word_count : dictionary
        number of words that appear in the function column.
    """
    word_count = {}
    #loop to count words
    for i in range(len(no_comments)):
        #convert a sentence into a list of words and remove english stop words
        text_tokens = word_tokenize(no_comments[i][3])
        word_list = [word for word in text_tokens if not word in stopwords.words('english')]
        #loop through each word from the sentence
        for word in word_list:
            #if new words are obtained
            if word not in word_count:
                #add the word to the dictionary and add first apperance together with the group and user
                word_count[word] = [0,no_comments[i][2],no_comments[i][4]]
            #if it is not a new word then increment the occurance by 1
            word_count[word][0]+= 1 
    return word_count


# sort words
def sort_words(word_count):
    """   

    Parameters
    ----------
    word_count : dictionary
        words and the number of times they occur.

    Returns
    -------
    Sorted List.

    """    
    sorted_word_frequency = []
    #sort the words in descendng order in the number of times the word appears
    for key, val in word_count.items():
        sorted_word_frequency.append((key, val))
    s = sorted(sorted_word_frequency, key = lambda x: x[1], reverse=True)
    return s

#create a turple
def create_tuple(list_of_sorted_words):
    """
    Parameters
    ----------
    list_of_sorted_words : list
        List of words that are sorted.
    Returns
    -------
    words_tuple_list : list
        returns a list that contains a turple of number of word, frequencies, Group and User
    """
    words_tuple_list= []
    for w in list_of_sorted_words:
        my_tup = (w[0], w[1][0], w[1][1], w[1][2])
        words_tuple_list.append(my_tup)
    return words_tuple_list

def list_to_dataframe(my_list):
    """
    Parameters
    ----------
    my_list : TYPE
        a list of records.
    Returns
    -------
    top_100_words : dataframe
        dataframe of top 100 words.
    """
    #convert data to dataframe
    df = pd.DataFrame(my_list, columns=['Word','Frequency','Organization','User'])
    #get the first 100 words
    top_100_words = df[:100]   
    
    return top_100_words


def create_piechart(top_100_words):
    """
    Parameters
    ----------
    top_100_words : datafarame
        get a list of turples.

    Returns
    -------
    None.

    """
    #group the words by their user and find the number of words from that user
    group_by_user = top_100_words.groupby(['User']).sum()
    #make the piechart a perfect circle
    plt.axis("equal")
    #plot a piechart of frequency of words and user 
    plt.pie(group_by_user['Frequency'], labels=group_by_user.index, radius=2, autopct='%0.2f%%')
    plt.show()

def write_to_excel(data,excel_file,sheet_name):
    """   

    Parameters
    ----------
    data : dataframe
        data tha is to be written in excel.
    excel_file : string
        Name of the excel file to write to
    sheet_name : Sheet name to write to

    Returns
    -------
    None.

    """
    #write to excel file
    data.to_excel(excel_file,sheet_name=sheet_name, startrow=0, startcol=0,index=False)
    
def calculate_tf_idf(data_set,N):
    """
    Parameters
    ----------
    data_set : list
        List of tuples of the dataset.
    N : Integer
        Number of documnents( users or Organization for calculating df-idf)        

    Returns
    -------
    Dataframe of column word, org/user, tf-idf representation of the dataset        

    """
    #will the tf_idf caluculations and the other data.
    word_tf_idf = []
    #loop through the data set to calculate tfidf 
    for w in data_set:
        #get tf(occurrance of word in the document)
        tf = w[1]
        try:
            #calculate idf = log(occurance of words/ Number of documents)
            #log of 0 is undefined, so assign 0 if reach where you get log of 0
            idf = math.log((tf/N))
        except:
            idf = 0
   
        #tf-idf = tf * idf
        tf_idf = tf * idf
        #there are 3 org and 4 users, generate dataframe of specific group and return the data.
        if N == 4:            
            #add data to list
            tf_idf_data = (w[0],w[3],tf_idf)
        if N == 3:
            tf_idf_data = (w[0],w[2],tf_idf)
        #add tuple to list
        word_tf_idf.append(tf_idf_data)
    if N == 4:
        user_tf_idf_df = pd.DataFrame(word_tf_idf, columns=['Word','User','TF-IDF'])
        return user_tf_idf_df
    if N == 3:
        org_tf_idf_df = pd.DataFrame(word_tf_idf, columns=['Word','Organization','TF-IDF'])            
        return org_tf_idf_df
    


if __name__ == "__main__":
    excel_file = "Book1.xlsx"
    sheet_number = "Sheet1"
    
    formated_lst = read_excel_and_format(excel_file, sheet_number)
    
    no_comments = formart_functionality(formated_lst)
    
    word_count =  count_words(no_comments)
    
    sorted_word_counts = sort_words(word_count)
    
    words_tuple = create_tuple(sorted_word_counts)
    
    top_100_words = list_to_dataframe(words_tuple)
    
    create_piechart(top_100_words)
    # write count of words to excel
    write_to_excel(top_100_words,"WordCount.xlsx","Count")
    
    user_tf_idf= calculate_tf_idf(words_tuple[:100],4)
    org_tf_idf= calculate_tf_idf(words_tuple[:100],3)
    #write tf-idf of user and organizarion
    write_to_excel(user_tf_idf,"TF-IDF-USER.xlsx","User")
    write_to_excel(org_tf_idf,"TF-IDF-ORG.xlsx","Organization")
    
    
    