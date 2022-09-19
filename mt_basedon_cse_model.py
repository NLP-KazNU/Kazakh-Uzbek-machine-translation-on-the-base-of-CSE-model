import re
import os
import xlrd, xlwt


#------------------------- WORD in SL = STEM + ENDING in SL -------------------------#

def splitting_by_words(text):
    result = re.findall("\w+\'?\\w+", text)
    return result


def sorting_endings(endings_file_name):
    endings_wb = xlrd.open_workbook(endings_file_name)
    endings_sh = endings_wb.sheet_by_index(0)
    endings = []
    for rownum in range(endings_sh.nrows-1):
        ending = endings_sh.cell(rownum+1,1).value
        if '\ufeff' or '-' in ending:
            ending = ending.replace('\ufeff', '')
            ending = ending.replace('-', '')
        endings.append(ending)

    sorted_endings = sorted(endings, key=len, reverse=True)
    ###print(sorted_endings)

    return sorted_endings


def stem(word, endings, stems_file_name):
    stems_wb = xlrd.open_workbook(stems_file_name)
    stems_sh = stems_wb.sheet_by_index(0)
    stems = []
    for rownum in range(stems_sh.nrows-1):
        stem = stems_sh.cell(rownum+1,0).value
        if '\ufeff' in stem:
            stem = stem.replace('\ufeff', '')
        stems.append(stem)

    stems_list = sorted(stems, key=len, reverse=True)
    ###print(stems_list)
    
    word_len = len(word)
    min_len_of_word = 2

    if word_len > min_len_of_word:
        n = word_len - min_len_of_word

        if word in stems_list:
            rez_stem = word
            rez_ending = ""

        else:
            i = n+1
            while i > 0:
                word_ending = word[word_len - (i-1):]
                stem = word[:word_len-len(word_ending)]                         
                for ending in endings:
                    if word_ending == ending:
                        if stem in stems_list:
                            rez_stem = stem
                            rez_ending = word_ending
                            i = 0
                if word_ending == '':
                    if stem in stems_list:
                        rez_stem = stem
                        rez_ending = word_ending
                        i = 0
                    else:
                        j = n+1
                        while j > 0:
                            word_ending = word[word_len - (j-1):]
                            stem = word[:word_len-len(word_ending)]
                            for ending in endings:
                                if word_ending == ending:
                                    rez_stem = stem
                                    rez_ending = word_ending
                                    j = 0
                                    i = 0
                            if word_ending == '':
                                rez_stem = word
                                rez_ending = word_ending
                                j = 0
                                i = 0
                            j = j-1
                i = i-1

    else:
        rez_stem = word
        rez_ending = ""
                        
    return rez_stem, rez_ending
    

def stemming(tfile_name, endings, stopwords_file_name, stems_file_name):
    text_file = open(tfile_name, 'r', encoding="utf-8")
    text_file = text_file.read()

    stopwords_wb = xlrd.open_workbook(stopwords_file_name)
    stopwords_sh = stopwords_wb.sheet_by_index(0)
    stopwords = []
    for rownum in range(stopwords_sh.nrows-1):
        stopword = stopwords_sh.cell(rownum+1,0).value
        if '\ufeff' in stopword:
            stopword = stopword.replace('\ufeff', '')
        stopwords.append(stopword)
    stopwords_list = sorted(stopwords, key=len, reverse=True)
    ###print(stopwords_list)
        
    text = splitting_by_words(text_file)
    ###print("text=",text)
    res_text = []

    rim_cifry = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x', 'xi', 'xii', 'xiii', 'xiv', 'xv', 'xvi', 'xvii', 'xviii', 'xix', 'xx', 'xxi', 'xxiv']

    for word in text:
        word = word.lower()
        if word not in res_text:
            if word.isnumeric() or word in rim_cifry:
                continue
            res_text.append(word)

    result_words  = [word for word in res_text if word not in stopwords_list]
    ##print("rez=",result_words)
    
    text_by_stemending = {}
    for word in result_words:
        stemm, endingg = stem(word, endings, stems_file_name)
        text_by_stemending.update({word: [stemm, endingg]})
    ##print(text_by_stemending)
       
    return text_by_stemending


#------------------------- STEM + ENDING in SL = STEM + ENDING in TL = WORD in TL -------------------------#

def match_by_table(endings_file, stems_file, stemsandendings_insl):
    stems_wb = xlrd.open_workbook(stems_file)
    stems_sh = stems_wb.sheet_by_index(0)
    stems = {}
    for rownum in range(stems_sh.nrows-1):
        stem_in_sl = stems_sh.cell(rownum+1,0).value
        stem_in_tl = stems_sh.cell(rownum+1,1).value
        stems.update({stem_in_sl: stem_in_tl})
    ###print("Stems ", stems)

    endings_wb = xlrd.open_workbook(endings_file)
    endings_sh = endings_wb.sheet_by_index(0)
    endings = {}
    for rownum in range(endings_sh.nrows-1):
        ending_in_sl = endings_sh.cell(rownum+1,1).value
        ending_in_sl = ending_in_sl.replace('-', '')
        ending_in_tl = endings_sh.cell(rownum+1,4).value
        endings.update({ending_in_sl: ending_in_tl})
    ###print("Endings ", endings)

    matched_words = {}
    for word_in_sl in stemsandendings_insl.keys():
        ###print("*****", word_in_sl, "*****")
        stem_in_sl = stemsandendings_insl[word_in_sl][0]
        ending_in_sl = stemsandendings_insl[word_in_sl][1]
        #print(stem_in_sl, ending_in_sl)
        if stem_in_sl in stems.keys():
            stem_in_tl = stems[stem_in_sl]
            ###print(stem_in_tl)
        else:
            stem_in_tl = stem_in_sl
            ###print(stem_in_tl)
        if ending_in_sl in endings.keys():
            ending_in_tl = endings[ending_in_sl]
            ###print(ending_in_tl)
        else:
            ending_in_tl = ending_in_sl
            ###print(ending_in_tl)
        word_in_tl = stem_in_tl + ending_in_tl
        matched_words.update({word_in_sl: word_in_tl})
        #print(word_in_sl, word_in_tl)

    return matched_words

def replace_sltotl(text_file, matchings, stopwords_file_name):
    stopwords_wb = xlrd.open_workbook(stopwords_file_name)
    stopwords_sh = stopwords_wb.sheet_by_index(0)
    stopwords = {}
    for rownum in range(stopwords_sh.nrows-1):
        stopword_in_sl = stopwords_sh.cell(rownum+1,0).value
        stopword_in_tl = stopwords_sh.cell(rownum+1,1).value
        stopwords.update({stopword_in_sl: stopword_in_tl})
    ###print(stopwords)

    text_file = open(text_file, 'r', encoding="utf-8")
    text_file = text_file.read()
    punctuations = ['.', ',', '?', '!', ':', ';', '(', ')', '"', '«', '»']
    for punctn in punctuations:
        text_file = text_file.replace(punctn, " "+punctn+" ")
    text_file = text_file.replace("  ", " ")
    print("TEXT in SL")
    print(text_file)
    
    for word in matchings:
        ###print(word)
        if f"{word} " in text_file.lower():
            ###print(matchings[word])
            text_file = text_file.lower().replace(f"{word} ", matchings[word]+" ")

    for stopword in stopwords:
        ###print(stopword)
        if f" {stopword} " in text_file.lower():
            ###print(stopwords[stopword])
            text_file = text_file.lower().replace(f" {stopword} ", " "+stopwords[stopword]+" ")
            
    print("\nTEXT in TL")
    translated_text = text_file
    
    return translated_text

stopwords_file_name = "qaz-uz-stopwords.xlsx"

endings_file_name = "qaz-uz-tab.xlsx"
endings = sorting_endings(endings_file_name)
###print(endings)

text_file_name = "text-qaz.txt"
stems_file_name = "qaz-uz-stems.xlsx"
text_by_stemending = stemming(text_file_name, endings, stopwords_file_name, stems_file_name)
###print(text_by_stemending)

matchings = match_by_table(endings_file_name, stems_file_name, text_by_stemending)
###print(matchings)

translated_text = replace_sltotl(text_file_name, matchings, stopwords_file_name)
print(translated_text)
