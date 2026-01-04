package main;

import lombok.SneakyThrows;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Collection;
import java.util.Comparator;
import java.util.List;

public class FileHelper {

    public static File[] folderReader(String filePath) {
        File file = new File(filePath);
        File[] files = file.listFiles();
        assert files != null;
        List<File> fileList = Arrays.asList(files);
        fileList.sort(Comparator.comparing(File::getName));
        return files;
    }

    public static List<File> extractExcel(String path) {
        File folder = new File(path);
        Collection<File> files = FileUtils.listFiles(folder, null, true);
        return files.stream()
                .filter(file -> file.getName().startsWith("外场试验队"))
                .filter(file -> file.getName().endsWith(".xls"))
                .sorted(Comparator.comparing(File::getName))
                .toList();
    }

    @SneakyThrows
    public static void copyFile(String path) {
        List<File> files = extractExcel(path);
        Path dest = Paths.get(path + "-");
        Files.createDirectories(dest);
        for (File file : files) {
            FileUtils.copyFile(file, new File(dest.toString(), file.getName()));
        }
        FileUtils.deleteDirectory(new File(path));
        dest.toFile().renameTo(new File(path));
    }
}
