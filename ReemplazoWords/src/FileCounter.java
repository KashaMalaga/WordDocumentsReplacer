import java.io.File;

public class FileCounter {
    @SuppressWarnings("ResultOfMethodCallIgnored")
    public static int countFiles(String directory, String typeFile) {
        final int[] count = {0};
        File dir = new File(directory);
        if (!dir.exists() || !dir.isDirectory()) {
            return 0;
        }
        dir.listFiles(file -> {
            if (file.isDirectory()) {
                count[0] += countFiles(file.getAbsolutePath(), typeFile);
            } else if (file.getName().endsWith(typeFile)) {
                count[0]++;
            }
            return false;
        });
        return count[0];
    }
}