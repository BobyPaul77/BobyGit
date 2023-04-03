package Pack;

import java.io.IOException;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		Excel e= new Excel();
		String s = e.readData(0,0);
		System.out.println(s);
		String s1 = e.readData(1,0);
		System.out.println(s1);
		String s2 = e.readData(2,0);
		System.out.println(s2);
			}

}
