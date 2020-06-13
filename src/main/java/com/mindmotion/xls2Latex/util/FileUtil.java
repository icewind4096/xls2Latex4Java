package com.mindmotion.xls2Latex.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

public class FileUtil {
    /**
     * 目录是否存在
     */
    public static boolean PathExists(String path) {
        return new File(path).exists();
    }

    /**
     * 文件是否存在
     */
    public static boolean FileExists(String path) {
        File file = new File(path);
        return file.exists() && file.isFile();
    }

    public static boolean saveToFileByList(String fileName, List<String> lists) {
        try {
            BufferedWriter bufferedWriter = new BufferedWriter(new FileWriter(fileName));
            for (String text: lists){
                bufferedWriter.write(text);
                bufferedWriter.newLine();
            }
            bufferedWriter.flush();
            bufferedWriter.close();
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }
}
