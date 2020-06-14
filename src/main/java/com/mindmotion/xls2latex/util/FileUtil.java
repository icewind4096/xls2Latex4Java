package com.mindmotion.xls2latex.util;

import org.apache.commons.io.FileUtils;

import java.io.File;
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
            FileUtils.writeLines(new File(fileName), lists);
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }
}
