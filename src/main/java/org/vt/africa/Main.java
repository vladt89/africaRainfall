package org.vt.africa;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

/**
 * @author vladimir.tikhomirov
 */
public class Main {

    private static final String DATA_FOLDER = "data";

    public static void main(String[] args) {
        FileReader fileReader = new FileReader();
        Stream<Path> walk = null;
        try {
            walk = Files.walk(Paths.get(DATA_FOLDER));
            walk.forEach(filePath -> {
                if (Files.isRegularFile(filePath)) {
                    System.out.println("\nGoing to parse following file: " + filePath);
                    File file = new File(filePath.toUri());
                    fileReader.fetchDataFromFile(file);
                }
            });
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (walk != null) {
                walk.close();
            }
        }
    }
}
