package com.l7.project1;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

public class TextToExcel {
	public static void main(String[] args) throws IOException {
		CreateExcel writer = new CreateExcel();
		StringBuilder sb =new StringBuilder();
		File file=new File("book.txt");
		Scanner sc =new Scanner(file);
		try {
				System.out.println("file found....");
				while(sc.hasNextLine()) {
					String test=sc.nextLine();
					sb.append(test);
					sb.append("\n");
				}
				System.out.println(sb);
				writer.createexcel(sb,"book_details.xlsx");
				sc.close();
				System.out.println("sucessfull");
		} catch (FileNotFoundException e) {
			sc.close();
		} catch (FormatError e) {
			System.out.println(e.getMessage());
		}	
	}

}
