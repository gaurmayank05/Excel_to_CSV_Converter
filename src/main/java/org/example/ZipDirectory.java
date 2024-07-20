package org.example;

import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ZipDirectory {

    /**
     * Creates a temporary directory with the specified prefix and returns its absolute path.
     *
     * @param fileName the prefix string to be used in generating the directory's name; may be {@code null}
     * @return the absolute path of the newly created temporary directory
     * @throws IOException if an I/O error occurs or the temporary-file directory does not exist
     */
    public String createTempDirectory (String fileName) throws IOException {
        return Files.createTempDirectory(fileName).toFile().getAbsolutePath();
    }

    /**
     *
     * @param tempDirectoryPath is used to store the path of file (or folder) which need to be deleted
     */
    public void deleteTempDirectory (String tempDirectoryPath) {
        File tempFile = new File(tempDirectoryPath);
        if (tempFile.exists()) {
            FileUtils.deleteQuietly(tempFile);
        }
    }

    /**
     *
     * @param sourceFolder is used to store the path of files (or folders) for which zip folder is created
     * @param destinationFolder is used to store the path where zip file need to be created
     * @throws IOException if an I/O error occurs or the temporary-file directory does not exist
     */
    public void zipFolder (String sourceFolder, String destinationFolder) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(destinationFolder);
        ZipOutputStream zipOut = new ZipOutputStream(fileOut)) {
            File fileToZip = new File(sourceFolder);
            File[] subFolder = fileToZip.listFiles();
            if (subFolder != null) {
                for (File childFolder : subFolder){
                    createZipFile(childFolder, childFolder.getName(), zipOut);
                }
            }
        }
    }

    /**
     * Adds a file or directory to the ZIP output stream.
     * This method recursively processes the specified file or directory and adds it to the
     * provided. If the file is hidden, it is skipped. If the file is a directory,
     * it is added as an entry and its contents are recursively added. If the file is a regular file,
     * its contents are written to the ZIP output stream.
     *
     * @param fileToZip the file or directory to add to the ZIP output stream
     * @param fileName the name to use for the file or directory within the ZIP archive
     * @param zipOut the ZIP output stream to write the file or directory to
     * @throws IOException if an I/O error occurs during reading the file or writing to the ZIP output stream
     */
    private void createZipFile (File fileToZip, String fileName, ZipOutputStream zipOut) throws IOException {
        if (fileToZip.isHidden()) {
            return;
        }
        if (fileToZip.isDirectory()) {
            if (fileName.endsWith("/")) {
                zipOut.putNextEntry(new ZipEntry(fileName));
                zipOut.closeEntry();
            }else{
                zipOut.putNextEntry(new ZipEntry(fileName + "/"));
                zipOut.closeEntry();
            }
            File[] children = fileToZip.listFiles();
            if (children != null) {
                for (File childFile : children) {
                    createZipFile(childFile, fileName + "/" + childFile.getName(), zipOut);
                }
            }
            return;
        }
        FileInputStream file = new FileInputStream(fileToZip);
        ZipEntry zipEntry = new ZipEntry(fileName);
        zipOut.putNextEntry(zipEntry);
        byte[] bytes = new byte[1024];
        int length;
        while ((length = file.read(bytes)) >= 0){
            zipOut.write(bytes, 0, length);
        }
        file.close();
    }
}