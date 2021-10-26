import com.gembox.spreadsheet.*;

import java.io.IOException;
import java.util.*;

class MostRepeatedWords {

    ArrayList<String> MRW = new ArrayList<>();

    MostRepeatedWords() {

        try {

            ExcelFile RW = ExcelFile.load("D:\\computer\\InformationRetrieval\\project\\mostly repeated words.xlsx");
            ExcelWorksheet worksheet = RW.getWorksheet(0);
            ExcelRow row;
            ExcelCell cell;

            for (int i = 0; i < 24; i++) {

                row = worksheet.getRow(i);
                cell = row.getCell(0);
                MRW.add(cell.getStringValue());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

class Posting {

    int docID;
    int tf;
    double tf_idf;

    ArrayList<Integer> position = new ArrayList<>();

    Posting(int docID, int tf) {

        this.docID = docID;
        this.tf = tf;
    }
}

class PostingsList {

    ArrayList<Posting> postingsList = new ArrayList<>();
    int count = 0;
    int df = 0;

    void AddPosting(int docID, int position) {

        boolean exist = false;

        for (Posting posting : postingsList) {

            if (posting.docID == docID) {

                postingsList.remove(posting);
                posting.tf++;
                posting.position.add(position);
                postingsList.add(posting);
                exist = true;
                break;
            }
        }

        if (!exist) {

            Posting posting = new Posting(docID, 1);
            posting.position.add(position);
            postingsList.add(posting);
            df++;
        }

        count++;
    }
}

class Dictionary {

    HashMap<String, PostingsList> dictionary = new HashMap<>();
    HashMap<String, Integer> word_index = new HashMap<>();

    int words = 0;
    int distinctWords;


    void Tokenize(String news, int docID, MostRepeatedWords RW) {
        int start = 0, end;
        String word;
        int position = 0;

        for (int i = 0; i < news.length(); i++) {

            int temp = (int) news.charAt(i);

            if ((temp < 1570) || (temp > 1740) || (i == news.length() - 1)) {

                end = i;

                word = news.substring(start, end);

                if (i == news.length() - 1)
                    word += news.substring(news.length() - 1);

                if ((word.length() > 0) && (!RW.MRW.contains(word))) {

                    words++;

                    if (dictionary.get(word) != null)
                        dictionary.get(word).AddPosting(docID, position);
                    else {
                        PostingsList postingsList = new PostingsList();
                        postingsList.AddPosting(docID, position);
                        dictionary.put(word, postingsList);
                    }
                }

                position++;
                start = i + 1;
            }
        }
    }

    void UpdateDictionary() {

        int index = 0;
        distinctWords = 0;

        Iterator hmIterator = dictionary.entrySet().iterator();

        while (hmIterator.hasNext()) {

            Map.Entry mapElement = (Map.Entry) hmIterator.next();
            word_index.put((String) mapElement.getKey(), index);
            index++;
            distinctWords++;
        }
    }
}

class WordUnification {

    Dictionary dictionary;

    WordUnification(Dictionary dictionary) {
        this.dictionary = dictionary;
        modifyDictionary();
    }

    private PostingsList countBigrams(PostingsList pl1, PostingsList pl2) {

        PostingsList result = new PostingsList();

        for (Posting posting1 : pl1.postingsList) {

            for (Posting posting2 : pl2.postingsList) {

                if (posting2.docID > posting1.docID)
                    break;

                if (posting2.docID == posting1.docID) {

                    for (Integer p1 : posting1.position) {

                        for (Integer p2 : posting2.position) {

                            if (p2 > (p1 + 1))
                                break;

                            if (p2 == (p1 + 1))
                                result.AddPosting(posting1.docID, p1);
                        }
                    }
                }
            }
        }

        return result;
    }

    private PostingsList countTrigrams(PostingsList pl1, PostingsList pl2, PostingsList pl3) {

        PostingsList result = new PostingsList();

        for (Posting posting1 : pl1.postingsList) {

            for (Posting posting2 : pl2.postingsList) {

                if (posting2.docID > posting1.docID)
                    break;

                if (posting2.docID == posting1.docID) {

                    for (Posting posting3 : pl3.postingsList) {

                        if (posting3.docID > posting1.docID)
                            break;

                        if (posting3.docID == posting1.docID) {

                            for (Integer p1 : posting1.position) {

                                for (Integer p2 : posting2.position) {

                                    if (p2 > (p1 + 1))
                                        break;

                                    if (p2 == (p1 + 1)) {

                                        for (Integer p3 : posting3.position) {

                                            if (p3 > (p2 + 1))
                                                break;

                                            if (p3 == (p2 + 1))
                                                result.AddPosting(posting1.docID, p1);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        return result;
    }

    private double BackoffModel(PostingsList bigram, PostingsList trigram, PostingsList lastWord) {

        double PUnigram = ((double) lastWord.count) / ((double) dictionary.words);
        double PBigram = ((double) bigram.count) / ((double) lastWord.count);
        double PTrigram = ((double) trigram.count) / ((double) bigram.count);
        return (0.7 * PTrigram) + (0.2 * PBigram) + (0.1 * PUnigram);
    }

    private void modifyDictionary() {

        HashMap<String, PostingsList> to_be_add = new HashMap<>();

        dictionary.dictionary.forEach((k1, v1) -> {

            dictionary.dictionary.forEach((k2, v2) -> {


                PostingsList bigram = countBigrams(v1, v2);
                double PBigram = ((double) bigram.count) / ((double) v1.count);

                if (PBigram > 0.6)
                    to_be_add.put(k1 + k2, bigram);

                dictionary.dictionary.forEach((k3, v3) -> {

                    PostingsList trigram = countTrigrams(v1, v2, v3);

                    if (BackoffModel(bigram, trigram, v3) > 0.6)
                        to_be_add.put(k1 + k2 + k3, trigram);
                });
            });
        });

        to_be_add.forEach((k, v) -> {
            dictionary.dictionary.put(k, v);
        });
    }
}

class RootVerbs {

    private static double Threshold = 1;

    private ArrayList<String> RVL = new ArrayList<>();
    private Dictionary dictionary;

    RootVerbs(Dictionary dictionary) {

        this.dictionary = dictionary;

        try {
            ExcelFile RV = ExcelFile.load("D:\\computer\\InformationRetrieval\\project\\verbs.xlsx");
            ExcelWorksheet worksheet = RV.getWorksheet(0);
            ExcelRow row;
            ExcelCell cell;

            for (int i = 0; i < 48; i++) {

                row = worksheet.getRow(i);
                cell = row.getCell(0);
                RVL.add(cell.getStringValue());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        modifyDictionary();
    }

    private PostingsList merge(PostingsList pl1, PostingsList pl2) {

        PostingsList result = new PostingsList();

        int pointer = 0;

        for (Posting p1 : pl1.postingsList) {

            for (Posting p2 : pl2.postingsList) {

                if (p2.docID > pointer) {

                    if (p2.docID < p1.docID) {
                        result.AddPosting(p2.docID, 0);
                        pointer = p2.docID;

                    } else if (p2.docID == p1.docID) {
                        result.AddPosting(p2.docID, 0);
                        result.AddPosting(p2.docID, 0);
                        pointer = p2.docID;
                        break;

                    } else {
                        result.AddPosting(p1.docID, 0);
                        break;
                    }
                }
            }
        }

        return result;
    }

    void exchange(String word1, String word2, PostingsList postingsList) {

        dictionary.dictionary.remove(word1, postingsList);
        dictionary.dictionary.putIfAbsent(word2, postingsList);

        PostingsList postings;
        if (dictionary.dictionary.get(word2) != null) {
            postings = dictionary.dictionary.get(word2);
            postings = merge(postings, postingsList);
            dictionary.dictionary.put(word2, postings);
        }

    }

    private void modifyDictionary() {

        HashMap<String, String> change = new HashMap<>();
        HashMap<String, PostingsList> insert = new HashMap<>();

        dictionary.dictionary.forEach((k, v) -> {

            double idf = Math.log(7000 / ((double) v.df));

            if (idf < Threshold) {

                for (String root : RVL) {

                    if (k.contains(root)) {

                        change.put(k, root);
                        insert.put(root, v);
                        break;
                    }
                }
            }
        });

        change.forEach((k, v) -> {

            PostingsList postingsList = insert.get(v);
            exchange(k, v, postingsList);
        });
    }
}

class Addition extends RootVerbs {

    private HashMap<String, String> additions = new HashMap<>();
    Dictionary dictionary;

    Addition(Dictionary dictionary) {

        super(dictionary);

        this.dictionary = dictionary;

        try {
            ExcelFile addition = ExcelFile.load("D:\\computer\\InformationRetrieval\\project\\addition.xlsx");
            ExcelWorksheet worksheet = addition.getWorksheet(0);
            ExcelRow row;
            ExcelCell cell1;
            ExcelCell cell2;

            for (int i = 0; i < 42; i++) {

                row = worksheet.getRow(i);
                cell1 = row.getCell(0);
                cell2 = row.getCell(1);
                additions.put(cell1.getStringValue(), cell2.getStringValue());
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        modifyDictionary2();
    }

    private void modifyDictionary2() {

        HashMap<String, String> change = new HashMap<>();
        HashMap<String, PostingsList> insert = new HashMap<>();

        dictionary.dictionary.forEach((k1, v1) -> {

            additions.forEach((k2, v2) -> {

                if (k1.equals(v2)) {
                    change.put(k1, k2);
                    insert.put(k2, v1);
                }
            });
        });

        change.forEach((k, v) -> {

            PostingsList postingsList = insert.get(v);
            exchange(k, v, postingsList);
        });
    }
}

class Similarity {

    int docID;
    double similarity;
}

class Heap {

    private Similarity[] similarities;
    private int elements;

    Heap(Similarity[] similarities, int elements) {

        this.similarities = similarities;
        this.elements = elements;

        BuildMaxHeap();
    }

    private void Swap(int index1, int index2) {
        Similarity temp = similarities[index1];
        similarities[index1] = similarities[index2];
        similarities[index2] = temp;
    }

    private void MaxHeapify(int index) {

        int biggest;

        if ((((index * 2) + 1) < elements) && (((index * 2) + 2) < elements)) {

            if (similarities[index].similarity < similarities[(index * 2) + 1].similarity)
                biggest = (index * 2) + 1;
            else
                biggest = index;

            if (similarities[biggest].similarity < similarities[(index * 2) + 2].similarity)
                biggest = (index * 2) + 2;

            if (biggest != index) {
                Swap(index, biggest);
                MaxHeapify(biggest);
            }

        } else if (((index * 2) + 1) < elements) {

            if (similarities[index].similarity < similarities[(index * 2) + 1].similarity) {

                biggest = (index * 2) + 1;
                Swap(index, biggest);
                MaxHeapify(biggest);
            }

        } else if (((index * 2) + 2) < elements) {

            if (similarities[index].similarity < similarities[(index * 2) + 2].similarity) {

                biggest = (index * 2) + 2;
                Swap(index, biggest);
                MaxHeapify(biggest);
            }
        }
    }

    private void BuildMaxHeap() {

        for (int i = (elements - 1); i >= 0; i--)
            MaxHeapify(i);

    }

    int MaxHeap() {

        int result = similarities[0].docID + 1;

        similarities[0].similarity = -1;
        MaxHeapify(0);

        return result;
    }
}

class PostingsWeights {

    Dictionary dictionary;

    PostingsWeights(Dictionary dictionary) {

        this.dictionary = dictionary;

        ComputePostingsWeights();
    }

    private void ComputePostingsWeights() {

        dictionary.dictionary.forEach((k, v) -> {

            int df = v.df;

            for (Posting posting : v.postingsList) {

                int tf = posting.tf;

                posting.tf_idf = (1 + Math.log((double) tf)) * Math.log(7000 / (double) df);
            }
        });
    }
}

class ChampionLists {

    Dictionary dictionary;

    ChampionLists(Dictionary dictionary) {

        this.dictionary = dictionary;

        determineChampionLists();
    }

    private void determineChampionLists() {

        double Threshold = 0.5;

        PostingsWeights postingsWeights = new PostingsWeights(dictionary);
        dictionary = postingsWeights.dictionary;

        ArrayList<Posting> to_be_deleted = new ArrayList<>();

        Iterator hmIterator = dictionary.dictionary.entrySet().iterator();

        while (hmIterator.hasNext()) {

            Map.Entry mapElement = (Map.Entry) hmIterator.next();

            PostingsList v = (PostingsList) mapElement.getValue();

            for (Posting posting : v.postingsList) {

                if (posting.tf_idf < Threshold)
                    to_be_deleted.add(posting);
            }

            for (Posting posting : to_be_deleted) {

                v.postingsList.remove(posting);
                v.df--;
            }

            to_be_deleted = new ArrayList<>();
        }
    }
}

class Vector {

    private int size;

    Vector(int size) {
        this.size = size;
    }

    double VectorSize(double[] vector) {

        double result = 0;

        for (int i = 0; i < size; i++)
            result += Math.pow(vector[i], 2);

        result = Math.pow(result, 0.5);

        return result;
    }

    double VectorDot(double[] v1, double[] v2) {

        double result = 0;

        for (int i = 0; i < size; i++)
            result += v1[i] * v2[i];

        return result;
    }
}

class AnswerNoEL {

    private MostRepeatedWords repeatedWords;

    private Dictionary dictionary;
    private String query;

    private double[] QueryVector;
    private double[][] VectorSpace;

    private Similarity[] similarities;

    int[] results;

    AnswerNoEL(Dictionary dictionary, String query, MostRepeatedWords repeatedWords) {

        this.dictionary = dictionary;
        this.query = query;
        this.repeatedWords = repeatedWords;

        returnResults();
    }

    private void BuildVectorSpace() {

        VectorSpace = new double[7000][dictionary.distinctWords];

        dictionary.dictionary.forEach((k, v) -> {

            int index = dictionary.word_index.get(k);

            for (Posting posting : v.postingsList) {

                int tf = posting.tf;
                int df = v.df;
                double tf_idf = (1 + Math.log((double) tf)) * Math.log(7000 / ((double) df));
                VectorSpace[posting.docID - 1][index] = tf_idf;
            }
        });
    }

    private double Compute_tf_idf(String word, PostingsList postingsList) {

        int tf = postingsList.postingsList.get(0).tf;
        int df;
        if (this.dictionary.dictionary.get(word) == null)
            df = 7000;
        else
            df = this.dictionary.dictionary.get(word).df;

        return (1 + Math.log((double) tf)) * Math.log(7000 / ((double) df));
    }

    private void BuildQueryVector() {

        QueryVector = new double[dictionary.distinctWords];

        Dictionary dictionary = new Dictionary();

        dictionary.Tokenize(query, 0, repeatedWords);

        dictionary.UpdateDictionary();

        dictionary.dictionary.forEach((k, v) -> {

            double tf_idf = Compute_tf_idf(k, v);

            if (tf_idf > 0) {

                int index = this.dictionary.word_index.get(k);
                QueryVector[index] = tf_idf;
            }
        });
    }

    private void ComputeSimilarities() {

        similarities = new Similarity[7000];
        for (int i = 0; i < 7000; i++) {
            similarities[i] = new Similarity();
            similarities[i].docID = i;
        }

        BuildVectorSpace();
        BuildQueryVector();

        Vector vector = new Vector(dictionary.distinctWords);

        double queryVectorSize = vector.VectorSize(QueryVector);

        for (int i = 0; i < 7000; i++) {

            double vectorSize = vector.VectorSize(VectorSpace[i]);
            double dot = vector.VectorDot(VectorSpace[i], QueryVector);
            similarities[i].similarity = dot / (vectorSize * queryVectorSize);
        }
    }

    private void returnResults() {

        ComputeSimilarities();

        int K = 10;

        results = new int[K];

        Heap heap = new Heap(similarities, 7000);

        for (int i = 0; i < K; i++)
            results[i] = heap.MaxHeap();

    }
}

class AnswerEL {

    private MostRepeatedWords repeatedWords;

    private Dictionary dictionary;
    private String query;

    private double[] QueryVector;
    private Similarity[] similarities;

    private HashMap<Integer, double[]> VectorSpace = new HashMap<>();
    private int documents;
    private ArrayList<Integer> Documents = new ArrayList<>();

    static int K = 10;

    int[] results;

    AnswerEL(Dictionary dictionary, String query, MostRepeatedWords repeatedWords) {

        this.dictionary = dictionary;
        this.query = query;

        this.repeatedWords = repeatedWords;

        returnResults();
    }

    private Dictionary buildVectorSpace() {

        PostingsWeights pws = new PostingsWeights(this.dictionary);

        this.dictionary = pws.dictionary;

        Dictionary dictionary = new Dictionary();

        dictionary.Tokenize(query, 0, repeatedWords);

        dictionary.UpdateDictionary();

        dictionary.dictionary.forEach((k, v) -> {

            if (this.dictionary.dictionary.get(k) != null) {

                for (Posting posting : pws.dictionary.dictionary.get(k).postingsList) {

                    if (!Documents.contains(posting.docID)) {

                        Documents.add(posting.docID);
                        documents++;
                    }
                }
            }
        });

        Iterator hmIterator = this.dictionary.dictionary.entrySet().iterator();

        while (hmIterator.hasNext()) {

            Map.Entry mapElement = (Map.Entry) hmIterator.next();

            String k = (String) mapElement.getKey();
            PostingsList v = (PostingsList) mapElement.getValue();

            int df = v.df;

            for (Posting posting : v.postingsList) {

                if (Documents.contains(posting.docID)) {

                    int tf = posting.tf;

                    double tf_idf = (1 + Math.log((double) tf)) * Math.log(7000 / (double) df);

                    if (VectorSpace.get(posting.docID) == null) {

                        double[] vector = new double[this.dictionary.distinctWords];
                        vector[this.dictionary.word_index.get(k)] = tf_idf;

                        VectorSpace.put(posting.docID, vector);

                    } else {

                        double[] vector = VectorSpace.get(posting.docID);
                        vector[this.dictionary.word_index.get(k)] = tf_idf;

                        VectorSpace.put(posting.docID, vector);
                    }
                }
            }
        }

        return dictionary;
    }

    private void buildQueryVector(Dictionary dictionary) {

        QueryVector = new double[this.dictionary.distinctWords];

        dictionary.dictionary.forEach((k, v) -> {

            int tf = dictionary.dictionary.get(k).postingsList.get(0).tf;
            int df;
            if (this.dictionary.dictionary.get(k) == null)
                df = 7000;
            else
                df = this.dictionary.dictionary.get(k).df;

            double tf_idf = (1 + Math.log((double) tf)) * (Math.log(7000 / (double) df));

            if (tf_idf > 0) {

                int index = this.dictionary.word_index.get(k);
                QueryVector[index] = tf_idf;
            }
        });
    }

    private void ComputeSimilarities() {

        Dictionary dictionary = buildVectorSpace();

        buildQueryVector(dictionary);

        Vector vector = new Vector(this.dictionary.distinctWords);

        double QueryVectorSize = vector.VectorSize(QueryVector);

        similarities = new Similarity[documents];

        int i = 0;

        Iterator hmIterator = VectorSpace.entrySet().iterator();

        while (hmIterator.hasNext()) {

            Map.Entry mapElement = (Map.Entry) hmIterator.next();

            Similarity similarity = new Similarity();

            similarity.docID = (int) mapElement.getKey();

            double dot = vector.VectorDot((double[]) mapElement.getValue(), QueryVector);
            double VectorSize = vector.VectorSize((double[]) mapElement.getValue());

            similarity.similarity = dot / (VectorSize * QueryVectorSize);

            similarities[i] = similarity;
            i++;
        }
    }

    private void returnResults() {

        ComputeSimilarities();

        results = new int[K];

        Heap heap = new Heap(similarities, documents);

        for (int i = 0; i < K; i++)
            results[i] = heap.MaxHeap();
    }
}

class AnswerPhase1 {

    private Dictionary dictionary;
    private String query;
    private MostRepeatedWords mostRepeatedWords;
    private ArrayList<PostingsList> postingResults;
    private ArrayList<Integer> tempResults;
    int[] results;

    AnswerPhase1(Dictionary dictionary, String query, MostRepeatedWords repeatedWords) {

        this.dictionary = dictionary;
        this.query = query;
        this.mostRepeatedWords = repeatedWords;

        ComputeResults();
    }

    private PostingsList UnionPostings(PostingsList p1, PostingsList p2) {

        PostingsList result = new PostingsList();

        for (Posting posting1 : p1.postingsList) {

            for (Posting posting2 : p2.postingsList) {

                if (posting1.docID == posting2.docID) {

                    result.postingsList.add(posting1);
                    result.df++;
                }
            }
        }

        return result;
    }

    private void ComputeResults() {

        Dictionary dictionary = new Dictionary();

        dictionary.Tokenize(query, 0, mostRepeatedWords);

        ArrayList<PostingsList> postingsLists = new ArrayList<>();

        postingResults = new ArrayList<>();

        dictionary.dictionary.forEach((k, v) -> {

            if (this.dictionary.dictionary.get(k) != null) {

                postingsLists.add(this.dictionary.dictionary.get(k));
                System.out.println(k);
                postingResults.add(this.dictionary.dictionary.get(k));
            }
        });

        PostingsList unionPostings = postingsLists.get(0);

        for (PostingsList postingsList : postingsLists) {

            unionPostings = UnionPostings(unionPostings, postingsList);
            postingResults.add(unionPostings);
        }

        tempResults = new ArrayList<>();

        for (int i = postingResults.size() - 1; i >= 0; i--) {

            for (Posting posting : postingResults.get(i).postingsList) {

                if (!tempResults.contains(posting.docID))
                    tempResults.add(posting.docID);
            }
        }

        results = new int[tempResults.size()];

        for (int i = 0; i < tempResults.size(); i++)
            results[i] = tempResults.get(i);
    }
}

public class Main {

    private static int parse(String docID) {

        int result = 0;

        int length = docID.length();

        for (int i = 0; i < length; i++) {

            int temp = (int) docID.charAt(i);
            temp -= 48;
            temp *= Math.pow(10, length - i - 1);
            result += temp;
        }

        return result;
    }

    public static void main(String[] args) {

        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");
        SpreadsheetInfo.addFreeLimitReachedListener(eventArguments -> eventArguments.setFreeLimitReachedAction(FreeLimitReachedAction.CONTINUE_AS_TRIAL));

        MostRepeatedWords MRW = new MostRepeatedWords();

        try {

            ExcelFile dataset = ExcelFile.load("D:\\computer\\InformationRetrieval\\project\\IR_Spring2021_ph12_7k.xlsx");

            ExcelWorksheet worksheet = dataset.getWorksheet(0);
            ExcelRow row;
            ExcelCell cell;

            Dictionary dictionary = new Dictionary();
            String docID;
            String news;

            for (int i = 1; i <= 7000; i++) {

                row = worksheet.getRow(i);

                cell = row.getCell(0);
                docID = cell.getStringValue();

                cell = row.getCell(1);
                news = cell.getStringValue();

                int docId = parse(docID);
                if ((docId <= 7000) && (docId > 0))
                    dictionary.Tokenize(news, docId, MRW);
            }

//            WordUnification wordUnification = new WordUnification(dictionary);
//            dictionary = wordUnification.dictionary;
            Addition addition = new Addition(dictionary);
            dictionary = addition.dictionary;



            // phase 1
//            String query = "";

//            while (!query.equals("exit")) {
//
//                System.out.println("enter query");
//
//                Scanner reader = new Scanner(System.in);
//                query = reader.nextLine();
//
//                AnswerPhase1 answerPhase1 = new AnswerPhase1(dictionary, query, MRW);
//
//                String url;
//
//                for (int i = 0; i < 5; i++) {
//
//                    if (answerPhase1.results[i] <= 7000) {
//
//                        row = worksheet.getRow(answerPhase1.results[i]);
//                        cell = row.getCell(2);
//                        url = cell.getStringValue();
//
//                        System.out.println(answerPhase1.results[i] + "   " + url);
//                    }
//                }
//            }
            //





            ChampionLists championLists = new ChampionLists(dictionary);
            dictionary = championLists.dictionary;

            String query = "";

            while (!query.equals("exit")) {

                System.out.println("enter query");

                Scanner reader = new Scanner(System.in);
                query = reader.nextLine();

                dictionary.UpdateDictionary();

                //AnswerEL answerEL = new AnswerEL(dictionary, query, MRW);
                AnswerNoEL answerNoEL = new AnswerNoEL(dictionary, query, MRW);

                String url;

                for (int i = 0; i < AnswerEL.K; i++) {

                    if (answerNoEL.results[i] <= 7000) {

                        row = worksheet.getRow(answerNoEL.results[i]);
                        cell = row.getCell(2);
                        url = cell.getStringValue();

                        System.out.println(answerNoEL.results[i] + "   " + url);
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}