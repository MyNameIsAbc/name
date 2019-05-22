package excel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;

public class Utils {
	

	public static void getFileList(CopyOnWriteArrayList<File> filelist,String strPath) {
		File dir = new File(strPath);
		File[] files = dir.listFiles(); // 该文件目录下文件全部放入数组
		if (files != null) {
			for (int i = 0; i < files.length; i++) {
				String fileName = files[i].getName();
				if (files[i].isDirectory()) { // 判断是文件还是文件夹
					getFileList(filelist,files[i].getAbsolutePath()); // 获取文件绝对路径
				} else if (fileName.endsWith("xls")) { // 判断文件名是否以.avi结尾
					String strFileName = files[i].getAbsolutePath();
					System.out.println("-------" + strFileName);
					filelist.add(files[i]);
				} else {
					continue;
				}
			}

		}
	}

	// 打印日志
	static BufferedWriter bw;
	static FileWriter fileWriter;
	public static void print(String s, File file) {
		// 如果现在是使用FileOuputStream实例化，意味着所有的输出是向文件之中
	
		if (bw == null) {

			try {
				fileWriter=new FileWriter(file);
				bw = new BufferedWriter(fileWriter);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		Date ss = new Date();
		SimpleDateFormat format0 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String time = format0.format(ss.getTime());// 这个就是把时间戳经过处理得到期望格式的时间
		try {
			bw.write(time + " " + s);
			bw.flush();
			bw.newLine();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void stopprint() {
		try {
			
			bw.close();
			fileWriter.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
