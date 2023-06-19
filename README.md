# Search_engine
This is a simulated search engine implemented by Java.
The language of documents and queries for this search engine is Persian.

The project has been divided into four phases: 1- Preprocessing, 2- Document modeling in vector space, 3- Document clustering, 4- Query answering

1- Preprocessing:
The first step in this phase is tokenization. We tokenize documents and extract words. To do so, we consider spaces, commas, parentheses etc as separators.

The second step is text normalization. At this step, we unify some words that have come together consecutively many times and generate new tokens. In other words, we try to extract words that have multiple units. For instance, three small words will come together consecutively and carry a unique meaning (that happens a lot in Persian).
In order to do that, we used 2-gram and 3-gram language models.

The third step is eliminating uninformative words. For example, junctions, prepostionals etc.
This task is done by considering the frequency of each word in its relative document. Each word that has a frequency higher than a given threshold is considered uninformative and will be deleted.

The fourth step is finding the roots of words. There are many words in Persian that can come with different appearances in the case of time, number, singularity etc, but carry the same meaning. We determine which words have the same root (consequently the same meaning) and then replace them with their corresponding root.

At the end  of this phase, we save all the tokens in a dictionary.


2- Document modeling in vector space:
At the first step in this phase, we use the extracted tokens as inverted indexes for the documents.

Then, we use the tf-idf term to convert the documents to a vector space. Each dimension in the vector space determines an inverted index (token) and its corresponding value for each document is the tf-idf term of that token in the relative document.

3- Document clustering:
In this phase, we execute a clustering algorithm (K-means) on documents in the vector space. This will help us to answer queries faster.


4- Query answering:
At first, we receive the query from the user and will convert the query into the vector space in the same way as we did for documents.
For retrieving the proper documents, we first find the nearest cluster center to the query vector. Our distance measure for this task is Cosine Similarity.
Then, we calculate the Cosine Similarity of each document in the determined cluster to the query vector and insert them into a max-heap.
Then, we retrieve the K relevant documents (K highest Cosine Similarities) from the heap and display them on the console.

The first three phases are executed before asking queries from users. So, the execution time in those phases are not that much important.
However, the fourth phase should be executed in real time. That's why we don't sort all the documents and use a max-heap.
It is also worth it to mention that we have used index elimination and champion lists techniques to make the answering procedure faster.
Our search engine will answer the queries in 2 to 3 seconds.



Attention!!!!: the documents can not be uploaded to a git repository because their size is very high and it exceeds the upload size limit of the git. 
Only the code is available here. 
