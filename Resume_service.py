import shutil

import PyPDF2
import os
from io import StringIO
import pandas as pd
from collections import Counter
import en_core_web_sm

nlp = en_core_web_sm.load()
from spacy.matcher import PhraseMatcher
from flask_restful import Resource, Api, reqparse
import Paths

resumes_path = Paths.resumes_path
csv_path= Paths.csv_path
generated_excel_path=Paths.generated_excel_path
graph_path=Paths.graph_path
graph_path_png=Paths.graph_path_png
non_screenable_resumes=Paths.non_screenable_resumes
completed_screening=Paths.completed_screening

class ResumeScreen(Resource):

    def get(self):
        # Function to read resumes from the folder one by one
        mypath = resumes_path  # enter your path here where you saved the resumes
        onlyfiles = [os.path.join(mypath, f) for f in os.listdir(mypath) if os.path.isfile(os.path.join(mypath, f))]

        def pdfextract(file):
            fileReader = PyPDF2.PdfFileReader(open(file, 'rb'))
            countpage = fileReader.getNumPages()
            count = 0
            text = []
            while count < countpage:
                pageObj = fileReader.getPage(count)
                count += 1
                t = pageObj.extractText()
                print(t)
                text.append(t)
            return text

            # function to read resume ends

            # function that does phrase matching and builds a candidate profile

        def create_profile(file):
            text = pdfextract(file)
            text = str(text)
            text = text.replace("\\n", "")
            text = text.lower()
            # below is the csv where we have all the keywords, you can customize your own
            keyword_dict = pd.read_csv(csv_path)
            dotnet_assi = [nlp(text) for text in keyword_dict['.NetAssisted'].dropna(axis=0)]
            dotnet_dire = [nlp(text) for text in keyword_dict['.NetDirectProducts'].dropna(axis=0)]
            dotnet_sql = [nlp(text) for text in keyword_dict['.NetSDLLeadDeveloper'].dropna(axis=0)]
            datastage_mode = [nlp(text) for text in keyword_dict['Datastage+DataModelling'].dropna(axis=0)]
            datastage = [nlp(text) for text in keyword_dict['Datastage'].dropna(axis=0)]

            datastage_flink = [nlp(text) for text in keyword_dict['Datastage+Flink'].dropna(axis=0)]
            fullstack = [nlp(text) for text in keyword_dict['FullstackJava'].dropna(axis=0)]
            java_sql = [nlp(text) for text in keyword_dict['Java+SDLLead'].dropna(axis=0)]
            java = [nlp(text) for text in keyword_dict['Java'].dropna(axis=0)]
            java_backend = [nlp(text) for text in keyword_dict['JavaBackend'].dropna(axis=0)]

            java_frontend = [nlp(text) for text in keyword_dict['JavaFrontend'].dropna(axis=0)]
            devops = [nlp(text) for text in keyword_dict['DevOps'].dropna(axis=0)]
            etl_dev = [nlp(text) for text in keyword_dict['ETLDeveloper'].dropna(axis=0)]
            lsa = [nlp(text) for text in keyword_dict['LSA'].dropna(axis=0)]
            pega = [nlp(text) for text in keyword_dict['PegaCSSA'].dropna(axis=0)]

            ops_eng = [nlp(text) for text in keyword_dict['OpsEngineer'].dropna(axis=0)]
            pentester = [nlp(text) for text in keyword_dict['Pentester'].dropna(axis=0)]
            data_tester = [nlp(text) for text in keyword_dict['DataTester'].dropna(axis=0)]
            datamodel_dev = [nlp(text) for text in keyword_dict['DataModeler_DevEngineer'].dropna(axis=0)]
            datatester_dev = [nlp(text) for text in keyword_dict['DataTester_DevOps'].dropna(axis=0)]

            matcher = PhraseMatcher(nlp.vocab)
            matcher.add('.NetAssisted', None, *dotnet_assi)
            matcher.add('.NetDirectProducts', None, *dotnet_dire)
            matcher.add('.NetSDLLeadDeveloper', None, *dotnet_sql)
            matcher.add('Datastage+DataModelling', None, *datastage_mode)
            matcher.add('Datastage', None, *datastage)

            matcher.add('Datastage+Flink', None, *datastage_flink)
            matcher.add('FullstackJava', None, *fullstack)
            matcher.add('Java+SDLLead', None, *java_sql)
            matcher.add('Java', None, *java)
            matcher.add('JavaBackend', None, *java_backend)

            matcher.add('JavaFrontend', None, *java_frontend)
            matcher.add('DevOps', None, *devops)
            matcher.add('ETLDeveloper', None, *etl_dev)
            matcher.add('LSA', None, *lsa)
            matcher.add('PegaCSSA', None, *pega)

            matcher.add('OpsEngineer', None, *ops_eng)
            matcher.add('Pentester', None, *pentester)
            matcher.add('DataTester', None, *data_tester)
            matcher.add('DataModeler_DevEngineer', None, *datamodel_dev)
            matcher.add('DataTester_DevOps', None, *datatester_dev)
            doc = nlp(text)

            d = []
            matches = matcher(doc)
            for match_id, start, end in matches:
                rule_id = nlp.vocab.strings[match_id]  # get the unicode ID, i.e. 'COLOR'
                span = doc[start: end]  # get the matched slice of the doc
                d.append((rule_id, span.text))
            keywords = "\n".join(f'{i[0]} {i[1]} ({j})' for i, j in Counter(d).items())

            ## convertimg string of keywords to dataframe
            df = pd.read_csv(StringIO(keywords), names=['Keywords_List'])
            df1 = pd.DataFrame(df.Keywords_List.str.split(' ', 1).tolist(), columns=['Subject', 'Keyword'])
            df2 = pd.DataFrame(df1.Keyword.str.split('(', 1).tolist(), columns=['Keyword', 'Count'])
            df3 = pd.concat([df1['Subject'], df2['Keyword'], df2['Count']], axis=1)
            df3['Count'] = df3['Count'].apply(lambda x: x.rstrip(")"))

            base = os.path.basename(file)
            filename = os.path.splitext(base)[0]

            name = filename.split('_')
            name2 = name[0]
            name2 = name2.lower()
            ## converting str to dataframe
            name3 = pd.read_csv(StringIO(name2), names=['Candidate Name'])

            dataf = pd.concat([name3['Candidate Name'], df3['Subject'], df3['Keyword'], df3['Count']], axis=1)
            dataf['Candidate Name'].fillna(dataf['Candidate Name'].iloc[0], inplace=True)

            return (dataf)

        # function ends
        # code to execute/call the above functions
        try:
            final_database = pd.DataFrame()
            i = 0
            while i < len(onlyfiles):
                file = onlyfiles[i]
                try:
                    dat = create_profile(file)
                except Exception:
                    source = file
                    dest = non_screenable_resumes
                    try:
                        shutil.copy(source, dest)
                    except Exception:
                        print("file is already there")
                    print("something is not right")
                final_database = final_database.append(dat)
                if dat.shape[0]==1:
                    source = file
                    dest = non_screenable_resumes
                    try:
                        shutil.copy(source, dest)
                    except Exception:
                        print("file is already there")
                i += 1
                print(final_database)
        except Exception:
            source=file
            dest=non_screenable_resumes
            try:
                shutil.copy(source,dest)
            except Exception:
                print("file is already there")
            print("something is not right")

        # code to count words under each category and visulaize it through Matplotlib

        final_database2 = final_database['Keyword'].groupby(
            [final_database['Candidate Name'], final_database['Subject']]).count().unstack()
        final_database2.reset_index(inplace=True)
        final_database2.fillna(0, inplace=True)
        new_data = final_database2.iloc[:, 1:]
        new_data.index = final_database2['Candidate Name']
        json=new_data.to_json()
        # execute the below line if you want to see the candidate profile in a csv format
        sample2 = new_data.to_csv(generated_excel_path)
        import matplotlib.pyplot as plt

        plt.rcParams.update({'font.size': 10})
        ax = new_data.plot.barh(title="Resume keywords by category", legend=True, figsize=(25, 7), stacked=True)
        labels = []
        for j in new_data.columns:
            for i in new_data.index:
                label = str(int(new_data.loc[i][j]))
                labels.append(label)
        patches = ax.patches
        for label, rect in zip(labels, patches):
            width = rect.get_width()
            if width > 0:
                x = rect.get_x()
                y = rect.get_y()
                height = rect.get_height()
                ax.text(x + width / 2., y + height / 2., label, ha='center', va='center')
        plt.savefig(graph_path)
        plt.savefig(graph_path_png)
        tests = os.listdir(resumes_path)
        for item in tests:
            source_path=os.path.join(resumes_path, item)
            dest_path=completed_screening
            try:
                shutil.copy(source_path,dest_path)
            except Exception:
                print("file is already there")
            os.remove(os.path.join(resumes_path, item))
            print("file moved")
        return json





