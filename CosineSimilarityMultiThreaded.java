import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

// import java.time.Duration;
// import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.*;

public class CosineSimilarityMultiThreaded {

    private static Map<String, Map<String, Double>> distanceMatrix = new ConcurrentHashMap<>();

    public static void main(String[] args) throws IOException, InterruptedException {

        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss:SSS");

        Scanner scanner = new Scanner(System.in);

        System.out.println("Enter the path to the CSV file:");
        String csvPath = scanner.nextLine();

        System.out.println("Enter the number of threads to use:");
        int numThreads = scanner.nextInt();

        scanner.close();

        System.out.println(dtf.format(LocalDateTime.now()) + ": Reading File");

        List<PhraseVector> phraseVectors = readCSV(csvPath);

        System.out.println(dtf.format(LocalDateTime.now()) + ": Done reading File");

        System.out.println(dtf.format(LocalDateTime.now()) + ": Calculating distance matrix");
        ExecutorService executorService = Executors.newFixedThreadPool(numThreads);

        for (int i = 0; i < phraseVectors.size(); i++) {
            PhraseVector pv1 = phraseVectors.get(i);
            for (int j = i + 1; j < phraseVectors.size(); j++) {
                PhraseVector pv2 = phraseVectors.get(j);
                executorService.submit(() -> {
                    double distance = calculateCosineSimilarity(pv1.getPhrase(), pv2.getPhrase());
                    distanceMatrix.computeIfAbsent(pv1.getId(), k -> new ConcurrentHashMap<>()).put(pv2.getId(),
                            distance);
                    distanceMatrix.computeIfAbsent(pv2.getId(), k -> new ConcurrentHashMap<>()).put(pv1.getId(),
                            distance);
                });
            }
        }

        executorService.shutdown();
        executorService.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);

        System.out.println(dtf.format(LocalDateTime.now()) + ": Done calculating distance matrix");

        // printDistanceMatrix(distanceMatrix);

        System.out.println(dtf.format(LocalDateTime.now()) + ": Writing distance matrix to csv");

        String outfile = "P:\\Programming\\Java\\clustering\\OutputData.csv";
        writeDistanceMatrixToCsv(outfile, distanceMatrix);

        System.out.println(dtf.format(LocalDateTime.now()) + ": Processing complete");

    }

    public static List<PhraseVector> readCSV(String csvPath) throws IOException {
        List<PhraseVector> phraseVectors = new ArrayList<>();
        try (BufferedReader reader = new BufferedReader(new FileReader(csvPath))) {
            String line;
            while ((line = reader.readLine()) != null) {
                String[] tokens = line.split(";");
                // System.out.println("trying to parse line:\n" + line);
                if (tokens.length == 2) {
                    phraseVectors.add(new PhraseVector(tokens[0], tokens[1]));
                } else {
                    System.out.println("Could not parse line:\n" + line);
                }
            }
        }
        return phraseVectors;
    }

    public static double calculateCosineSimilarity(String phrase1, String phrase2) {
        Map<String, Integer> wordFrequency1 = getWordFrequency(phrase1);
        Map<String, Integer> wordFrequency2 = getWordFrequency(phrase2);

        Set<String> uniqueWords = new HashSet<>();
        uniqueWords.addAll(wordFrequency1.keySet());
        uniqueWords.addAll(wordFrequency2.keySet());

        double dotProduct = 0;
        double magnitude1 = 0;
        double magnitude2 = 0;

        for (String word : uniqueWords) {
            int freq1 = wordFrequency1.getOrDefault(word, 0);
            int freq2 = wordFrequency2.getOrDefault(word, 0);

            dotProduct += freq1 * freq2;
            magnitude1 += freq1 * freq1;
            magnitude2 += freq2 * freq2;
        }

        magnitude1 = Math.sqrt(magnitude1);
        magnitude2 = Math.sqrt(magnitude2);

        if (magnitude1 == 0 || magnitude2 == 0) {
            return 0;
        }

        double similarity = dotProduct / (magnitude1 * magnitude2);
        double distance = 1 - similarity;
        return distance;
    }

    public static Map<String, Integer> getWordFrequency(String phrase) {
        Map<String, Integer> wordFrequency = new HashMap<>();
        String[] words = phrase.split("\\s+");
        for (String word : words) {
            wordFrequency.put(word, wordFrequency.getOrDefault(word, 0) + 1);
        }
        return wordFrequency;
    }

    public static void printDistanceMatrix(Map<String, Map<String, Double>> distanceMatrix) {
        System.out.print("\t");
        for (String id : distanceMatrix.keySet()) {
            System.out.print(id + "\t");
        }
        System.out.println();
        for (String id1 : distanceMatrix.keySet()) {
            System.out.print(id1 + "\t");
            for (String id2 : distanceMatrix.keySet()) {
                System.out.printf("%.2f\t", distanceMatrix.get(id1).getOrDefault(id2, 0.0));
            }
            System.out.println();
        }
    }

    public static void writeDistanceMatrixToCsv(String filename, Map<String, Map<String, Double>> distanceMatrix) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filename))) {
            // Write the column headers
            writer.write(";");
            for (String id : distanceMatrix.keySet()) {
                writer.write(id + ";");
            }
            writer.newLine();

            // Write the distance matrix data
            for (String id1 : distanceMatrix.keySet()) {
                writer.write(id1 + ";");
                for (String id2 : distanceMatrix.keySet()) {
                    writer.write(String.format("%.2f;", distanceMatrix.get(id1).getOrDefault(id2, 0.0)));
                }
                writer.newLine();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
