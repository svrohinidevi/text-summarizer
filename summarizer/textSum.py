from flask import Flask, render_template, request
from docx import Document
from nltk.corpus import stopwords
from nltk.cluster.util import cosine_distance
import numpy as np
import networkx as nx
from os import path


class Summarize:
    def __init__(self, name):
        self.filename = name
        self.para = []
        self.result_data = ""

    def sentence_similarity(self, sent1, sent2, stop_words=None):
        if stop_words is None:
            stop_words = []

        sent1 = [w.lower() for w in sent1]
        sent2 = [w.lower() for w in sent2]

        all_words = list(set(sent1 + sent2))

        vector1 = [0] * len(all_words)
        vector2 = [0] * len(all_words)

        # build the vector for the first sentence
        for w in sent1:
            if w in stop_words:
                continue
            vector1[all_words.index(w)] += 1

        # build the vector for the second sentence
        for w in sent2:
            if w in stop_words:
                continue
            vector2[all_words.index(w)] += 1
        return 1 - cosine_distance(vector1, vector2)

    def build_similarity_matrix(self, sentences, stop_words):
        # Create an empty similarity matrix
        similarity_matrix = np.zeros((len(sentences), len(sentences)))

        for idx1 in range(len(sentences)):
            for idx2 in range(len(sentences)):
                if idx1 == idx2:  # ignore if both are same sentences
                    continue
                similarity_matrix[idx1][idx2] = self.sentence_similarity(sentences[idx1], sentences[idx2], stop_words)
        return similarity_matrix

    def generate_summary(self, para, n):
        stop_words = stopwords.words('english')
        summarize_text = []
        article = []
        for i in range(0, len(para)):
            data = para[i].split(". ")
            for j in range(0, len(data)):
                article.append(data[j])
        for sen in article:
            if sen == '':
                article.remove('')
        sentences = []

        for sentence in article:
            sentences.append(sentence.replace("[^a-zA-Z0-9]", " ").split(" "))
        if len(sentences) != 0:
            sentences.pop()
        print(sentences)
        sentence_similarity_matrix = self.build_similarity_matrix(sentences, stop_words)
        sentence_similarity_graph = nx.from_numpy_array(sentence_similarity_matrix)
        scores = nx.pagerank(sentence_similarity_graph)

        ranked_sentence = sorted(((scores[i], s) for i, s in enumerate(sentences)), reverse=True)
        if len(ranked_sentence) == 0:
            for sen in article:
                summarize_text.append(" ".join(sen))
        else:
            for i in range(n):
                summarize_text.append(" ".join(ranked_sentence[i][1]))
        print(summarize_text)
        self.result_data += ". ".join(summarize_text) + ".\n\n"

    def readpara(self):
        if path.exists(self.filename):
            document = Document(self.filename)
        else:
            self.result_data = '\nFile does not exist. Enter file name again.'
            return
        #print(document)
        i1 = 0
        i2 = 0
        for paragraph in document.paragraphs:
            if i1 == 0:
                for run in paragraph.runs:
                    if run.bold:
                        i1 = 1
                        i2 = 1
                        self.result_data += paragraph.text + "\n\n"
                        self.para = []
                        break
            if i2 == 1:
                i2 = 0
            else:
                if i1 == 1 and i2 == 0:
                    for run in paragraph.runs:
                        if run.bold:
                            self.generate_summary(self.para, int(len(self.para)))
                            self.result_data += paragraph.text + "\n\n"
                            self.para = []
                            break
                        else:
                            self.para.append(paragraph.text)
                            break
        self.generate_summary(self.para, int(len(self.para)))

    def getsummary(self):
        self.readpara()
        return self.result_data


class Webpage:
    def __init__(self):
        self.app = Flask(__name__)
        @self.app.route('/')
        def index():
            return render_template("summarizer.html")

        @self.app.route('/', methods=['POST'])
        def getfile():
            self.name = request.form['pathname']
            summary = Summarize(self.name)
            self.result_data = summary.getsummary()
            return render_template('pass.html', n=self.result_data, a=self.name)

        @self.app.route('/home')
        def homepage():
            return render_template("home.html")

        @self.app.route('/faq')
        def faqpage():
            return render_template("faq.html")

        @self.app.route('/contact')
        def contactpage():
            return render_template("contact.html")


if __name__ == "__main__":
    web = Webpage()
    web.app.run(debug=True)
