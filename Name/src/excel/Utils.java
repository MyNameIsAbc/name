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
		File[] files = dir.listFiles(); // ���ļ�Ŀ¼���ļ�ȫ����������
		if (files != null) {
			for (int i = 0; i < files.length; i++) {
				String fileName = files[i].getName();
				if (files[i].isDirectory()) { // �ж����ļ������ļ���
					getFileList(filelist,files[i].getAbsolutePath()); // ��ȡ�ļ�����·��
				} else if (fileName.endsWith("xls")) { // �ж��ļ����Ƿ���.avi��β
					String strFileName = files[i].getAbsolutePath();
					System.out.println("-------" + strFileName);
					filelist.add(files[i]);
				} else {
					continue;
				}
			}

		}
	}

	// ��ӡ��־
	static BufferedWriter bw;
	static FileWriter fileWriter;
	public static void print(String s, File file) {
		// ���������ʹ��FileOuputStreamʵ��������ζ�����е���������ļ�֮��
	
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
		String time = format0.format(ss.getTime());// ������ǰ�ʱ�����������õ�������ʽ��ʱ��
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
