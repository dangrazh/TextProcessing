import java.util.HashMap;
import java.util.Map;

public class PhraseVector {

    private final String id;
    private final String phrase;
    private final Map<String, Integer> wordFrequency;

    public PhraseVector(String id, String phrase, Map<String, Integer> wordFrequency) {
        this.id = id;
        this.phrase = phrase;
        this.wordFrequency = wordFrequency;
    }

    public PhraseVector(String id, String phrase) {
        this.id = id;
        this.phrase = phrase;
        this.wordFrequency = this.getWordFrequency(phrase);
    }

    public String getId() {
        return id;
    }

    public String getPhrase() {
        return phrase;
    }

    public double cosineSimilarity(PhraseVector other) {
        double dotProduct = 0.0;
        double magnitude1 = 0.0;
        double magnitude2 = 0.0;
        for (String word : wordFrequency.keySet()) {
            int frequency1 = wordFrequency.get(word);
            int frequency2 = other.wordFrequency.getOrDefault(word, 0);
            dotProduct += frequency1 * frequency2;
            magnitude1 += frequency1 * frequency1;
        }
        for (String word : other.wordFrequency.keySet()) {
            int frequency2 = other.wordFrequency.get(word);
            magnitude2 += frequency2 * frequency2;
        }
        if (magnitude1 == 0.0 || magnitude2 == 0.0) {
            return 0.0;
        } else {
            return dotProduct / (Math.sqrt(magnitude1) * Math.sqrt(magnitude2));
        }
    }

    private Map<String, Integer> getWordFrequency(String phrase) {
        String[] words = phrase.toLowerCase().split("[^a-zA-Z0-9']+");
        Map<String, Integer> wordFrequency = new HashMap<>();
        for (String word : words) {
            int count = wordFrequency.getOrDefault(word, 0);
            wordFrequency.put(word, count + 1);
        }
        return wordFrequency;
    }
}
