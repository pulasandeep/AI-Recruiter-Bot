import docx2txt
import xlsxwriter
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity

resumes = ["JohnnyDeppResume.docx","JohnWickResume.docx","MannyJohnsResume.docx","TimothyFlowersResume.docx"]
candidates=["JohnnyDeppResume","JohnWickResume","MannyJohnsResume","TimothyFlowersResume"]
outWorkbook = xlsxwriter.Workbook("scores.xlsx")
outsheet = outWorkbook.add_worksheet("scores")
outsheet.write("A1","Candidate Name")
outsheet.write("B1","Resume Score")
alphas ="AB"
names = []
scrs = []
for i in range(len(resumes)):
    names.append(alphas[0]+str(i+2))
    scrs.append(alphas[1]+str(i+2))
print("Similarity scores of each candidates are listed below:")
scores={}
for member in range(len(resumes)):
    resume = docx2txt.process(resumes[member])
    job_description = docx2txt.process("Web Developer Job Description.docx")
    text = [resume, job_description]
    count_vector= CountVectorizer()
    count_matrix = count_vector.fit_transform(text)
    #print("\n Similarity Scores:")
    #print(cosine_similarity(count_matrix))
    match_percentage = cosine_similarity(count_matrix)[0][1]*100
    match_percentage= round(match_percentage,2)
    scores[candidates[member]]=match_percentage
    outsheet.write(names[member],candidates[member])
    outsheet.write(scrs[member],match_percentage)

outWorkbook.close()
