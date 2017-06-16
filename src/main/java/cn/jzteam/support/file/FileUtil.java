package cn.jzteam.support.file;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
/**
 * 文件操作
 */
public class FileUtil {
    
    public static void main(String[] args) {

        rename("D:\\jzteam\\music\\cache", ".mqcc", ".mp3");
        System.out.println("结束了");
    }

    /**
     * 
     * 描述：修改文件名，不改文件位置（把文件名中的srcNameSub替换成descNameSub）
     * 
     * @param filePath
     * @param srcNameSub
     * @param descNameSub
     * @return void
     * @exception
     * @createTime：2017年1月20日
     * @author: zhujz
     */
    public static void rename(String filePath, String srcNameSub, String descNameSub) {
        if (srcNameSub == null || srcNameSub.length() == 0) {
            return;
        }

        File file = new File(filePath);
        if (!file.exists()) {
            return;
        }

        fileRename(file, srcNameSub, descNameSub);
    }

    // 递归修改文件名
    private static void fileRename(File file, String srcNameSub, String descNameSub) {
        if (file == null || !file.exists()) {
            return;
        }

        if (file.isDirectory()) {
            File[] listFiles = file.listFiles();
            if (listFiles != null && listFiles.length > 0) {
                for (File f : listFiles) {
                    fileRename(f, srcNameSub, descNameSub);
                }
            }
        } else {
            String realPath = file.getPath();
            System.out.println("原路径:" + realPath);

            String parentPath = realPath.substring(0, realPath.lastIndexOf(File.separator));
            String fileName = file.getName().replace(srcNameSub, descNameSub);

            String filePath = parentPath + File.separator + fileName;
            System.out.println("新路径:" + filePath);

            file.renameTo(new File(filePath));
        }
    }

    private static void copyFile(File srcFile, File descFile) throws IOException {
        if (srcFile == null || descFile == null) {
            return;
        }

        if (!descFile.getParentFile().exists()) {
            descFile.getParentFile().mkdirs();
        }
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            bis = new BufferedInputStream(new FileInputStream(srcFile));
            bos = new BufferedOutputStream(new FileOutputStream(descFile));
            byte[] buf = new byte[1024];
            int len = 0;
            while ((len = bis.read(buf)) > 0) {
                bos.write(buf, 0, len);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            if (bis != null)
                bis.close();

            if (bos != null)
                bos.close();
        }
    }
}
