# -*- coding: utf-8 -*-
"""
Created on Tue Jun 14 19:00:35 2022

@author: debayand
"""

# import os
# import docx
import pdf2image
from PIL import Image
import cv2
import numpy as np
import pytesseract
import pkg_resources

pytesseract.pytesseract.tesseract_cmd='C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# pytesseract version
print(pkg_resources.working_set.by_key['pytesseract'].version)
# opencv version
print(cv2.__version__)

# convert pdf 2 image function
def convert_pdf_to_image(document,dpi):
    images=[]
    images.extend(list(map(lambda image:cv2.cvtColor(np.asarray(image),code=cv2.COLOR_BGR2GRAY),pdf2image.convert_from_path(document,dpi))))
    return images

images=convert_pdf_to_image("FW Ultron  Pre-MELT BOM Review.pdf", 300)
for i in range(len(images)):
    image=images[i]
    im=Image.fromarray(image)
    im.save("images/output_{}.jpg".format(i), "JPEG")
    each_page_text=pytesseract.image_to_string(im)
    print(each_page_text)
    # with open("out_%i.txt" %i, "w") as f:
    #     f.write(each_page_text)
    #     f.close()


# # Docx to String
# dock=docx.Document(r"PNR_Tool_Work_Instruction_V04.docx")
# data=""
# fulltext=[]
# for para in dock.paragraphs:
#     fulltext.append(para.text)
#     # fulltext = list(filter(None, fulltext))
#     data='\n'.join(fulltext)
    

# # print(data)

# dock2=docx.Document(r"SPEC-9457-PCN-PDN Process_Rev03.docx")
# data2=""
# fulltext2=[]
# for para in dock2.paragraphs:
#     fulltext2.append(para.text)
#     # fulltext = list(filter(None, fulltext))
#     data2='\n'.join(fulltext2)
    

# print(data2)

# import nltk
# import string
# # nltk.download_gui()
# from nltk import sent_tokenize
# from nltk import word_tokenize
# from nltk.corpus import stopwords
# from nltk import pos_tag
# from nltk.stem.lancaster import LancasterStemmer
# from nltk.stem.porter import PorterStemmer
# from nltk.stem.snowball import SnowballStemmer
# from nltk.stem import WordNetLemmatizer

# TEXT='''SINGLE USER LICENSE - NOT FOR USE ON A NETWORK OR ONLINE

# IPC-9592B November 2012

# APPENDIX E

# Manufacturing Reliability Testing - BI/HASS/ORT

# E.1 Introduction Manufacturing tests are typically conducted at room ambient (~25 °C) and are often only a few seconds in duration. These tests are useful for identifying obvious manufacturing and component defects that are easily detected and corrected. Additional testing may be required to identify product weaknesses that pass initial 25 °C testing but fail under more stressful conditions such as at temperature extremes or during the first few hours of operation.

# Failures that occur early in the life of a product are referred to as infant mortality. If undetected in factory testing, these failures may occur in the first year and often within the first month of operation in the field. While it is not practical to test for periods equivalent to one year in a manufacturing test, it is possible to test under accelerated stress conditions that allow detection of failures with just a few hours of exposure to those stresses.

# E.1.1 Overview of Test Options Manufacturing tests shall be put in place that apply stresses beyond 25 °C and for durations greater than those seen during standard functional test. Historically companies have used Burn-In (BI) to serve this purpose. BI is typically done at elevated temperatures (40 °C to 50 °C) while product is powered, loaded and periodically monitored for functionality. BI test duration may range from one hour to several days.

# An alternate method to accelerate failures is Highly Accelerated Stress Screening (HASS) and HASS Audit (HASA). This approach is more effective at identifying product weaknesses because it uses additional stressful environments such as thermal cycling and vibration, as well as input voltage (line) and load variation. While HASS is more effective at identifying infant mortality failures, it is also more difficult and expensive to implement in manufacturing due to added complexities of environmental chambers, test equipment and fixturing. However, due to the value HASS offers, it is strongly recommended that manufacturers have some HASS capability. Other tests of longer duration may be required on a sample basis. These tests are referred to as Ongoing Reliability Testing (ORT).

# E.1.2 Reliability Test Implementation - Best Practices Manufacturing reliability tests should not be viewed as effective screens for poor reliability. It is rare to be able to screen out all infant mortality failures. Failures in Burn-in, HASS and ORT should be viewed as early warnings of possible similar undetected failure mechanisms in a population of products. Failures representative of a given failure mechanism are often distributed in stress and/or time and may be undetectable with the stress-time limitations seen in production reliability tests. The most effective way to eliminate these failures is to rapidly diagnose the root cause of any failure identified during reliability testing and correct it.

# For this reason, it is strongly recommended that Power Supply manufacturers have the ability to perform rapid detailed failure analysis including same day access to laboratory facilities that support materials failure analysis (cross section, decap, SEM/EDX, etc.). These resources should prioritize any reliability test failures ahead of standard production yield improvement activities. In addition, it is recommended that manufacturers create an internal Failure Review Board (FRB) to review all failures for accurate root cause analysis and effectiveness of corrective action.

# As products mature, failures identified during manufacturing reliability tests should be rare. At some point it is no longer cost effective to continue performing reliability tests on 100% of manufactured products. The processes described in the following sections require products to start out with a 100% BI or HASS but allow for reduction in test time and a migration to sampling when failure rates are demonstrated to be low or zero, except in cases where parametric stability is required through BI/HASS on 100% of Product.

# It should be noted that HASS is effective at identifying more failure modes than traditional BI, however, in some cases it also may overstate the absolute risk of field failure. In other words, for some failure mechanisms HASS may identify failures with a low risk of appearing in the field. For example, HASS may identify a failure that only occurs with a power cycle between -20 °C and -40 °C and does not degrade with time. While this condition may be within specification, its probability of occurrence in normal field operation is low. Most failure mechanisms seen in HASS will eventually show up in the field, just possibly at a lower rate.

# When field failures indicate a reliability issue that has escaped detection in the factory, a frequent reaction is to increase burn-in time or HASS stresses. Just like the goal of all manufacturing tests should be to use test results to identify and correct defects, the same is true for field failures. The objective should be to identify the root cause of the failure and correct it, not to add a screen which may only be partially effective at best.

# '''
# # sentences = sent_tokenize(TEXT)
# words = word_tokenize(TEXT)
# # print(len(sentences))
# # print(len(words))
# stop_words=set(stopwords.words('english'))
# stop_words=stop_words.union(string.punctuation)

# cleanWords = [w for w in words if not w in stop_words]
# print(len(cleanWords))

# # for i in range(len(cleanWords)):
# #     print(str(i) + '---    ' + cleanWords[i])

# taggedWords = pos_tag(cleanWords)

# # for j in range(len(taggedWords)):
# #     print(taggedWords[j])

# # stemmer = LancasterStemmer()
# # stemmer = PorterStemmer()
# stemmer = SnowballStemmer('english')
# lemmatizer = WordNetLemmatizer()

# #compare the changes in cleanwords
# for word in cleanWords:
#     print(word + '  ---  '+ stemmer.stem(word) + '  --  ' + lemmatizer.lemmatize(word, pos='v'))
    