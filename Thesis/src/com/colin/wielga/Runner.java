package com.colin.wielga;

import java.awt.List;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Scanner;

public class Runner {

	public static boolean debug = false;
	public static ArrayList<Integer> lens = new ArrayList<Integer>();
	public static ArrayList<String> result = new ArrayList<String>();
	public static HashMap<String,Integer> counts = new HashMap<String,Integer>();

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		Counter.load("C:\\Users\\Colin\\Documents\\School\\Thesis 2\\encoding1.txt");
		LineEncoder.load("lineencoding1");
		Scanner s = new Scanner(System.in);
		
//		String[] temp = LineEncoder.encode("VBcode\\original\\VBProjectsFall03\\aahumann\\frmPick.frm");
//		for (int i =0;i<temp.length;i++){
//			System.out.println(temp[i]);
//		}
		
//		allEncodeLine("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
//		cmpall();
		// allPerceeds("C://Users//Colin//Documents//School//Thesis 2//VBCode//original","i",1);
		// Counter.printSet();

		// String[] temp =
		// Counter.encode(Counter.flaten("C://Users//Colin//Documents//School//Thesis 2//VBCode//original//VBProjectsFall03//aahumann//frmPick.frm"));
		// for (int i=0;i<temp.length;i++){
		// System.out.print(temp[i] +",");
		// }

		//System.out.println("" + Cmp.qc_rapper("acasdqaseasdqqafeadfawadasdveacasxczxcvdasdfeasdveghyagasdfvcrascvxawcefradvdacxvascvcxavzxcvawasdaqsasdfadasdfasdwdedasasdadce", "eadasdffasqasdasfrdedwsadsfascasdcadsfaedcvascqacasdvaeacvadsvasdveadfeadsfasdfasdsdasdfbhbhbfasdfedssqaswdesdassdwwwsdasdwfedasesfwaeqdswwasadcada"));

		// countEncoded("C://Users//Colin//Documents//School//Thesis 2//VBCode//original//VBProjectsFall03//CsciStudent");

		while (true) {
			System.out.println("ready for input");
			if (s.hasNext()) {
				String line = s.nextLine();
				if (line.startsWith("c1")) {
					countContent(
							"C://Users//Colin//Documents//School//Thesis 2//VBCode//original",
							1);
					Counter.getTotals();
					Counter.c = new HashMap<String, Integer>();
					System.out.println("done");
				} else if (line.startsWith("c2")) {
					countContent(
							"C://Users//Colin//Documents//School//Thesis 2//VBCode//original",
							2);
					Counter.getTotals();
					System.out.println("done");
				} else if (line.startsWith("c3")) {
					countContent(
							"C://Users//Colin//Documents//School//Thesis 2//VBCode//original",
							3);
					Counter.getTotals();
					Counter.c = new HashMap<String, Integer>();
					System.out.println("done");
				} else if (line.startsWith("c4")) {
					countContent(
							"C://Users//Colin//Documents//School//Thesis 2//VBCode//original",
							4);
					Counter.getTotals();
					Counter.c = new HashMap<String, Integer>();
					System.out.println("done");
				} else if (line.startsWith("c5")) {
					countContent(
							"C://Users//Colin//Documents//School//Thesis 2//VBCode//original",
							5);
					Counter.getTotals();
					Counter.c = new HashMap<String, Integer>();
					System.out.println("done");
				} else if (line.startsWith("ca")) {
					allcountcontent("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
					Counter.getTotals();
					Counter.c = new HashMap<String, Integer>();
					System.out.println("done");
				} else if (line.startsWith("ce")) {
					countEncoded("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
					// Counter.getTotals();
					Counter.getChecked();
					Counter.c = new HashMap<String, Integer>();
					System.out.println("done");
				} else if (line.startsWith("le")) {
					allEncodeLine("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
					// Counter.getTotals();
					System.out.println("done");
				} else if (line.startsWith("lle")) {
					llerapper("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
					// Counter.getTotals();
					System.out.println("done");
				} else if (line.startsWith("do it!")) {
					Runner.showMeTheMoney("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
					System.out.println("done");
				} else {
					System.out.println("unknown command");
				}

			}
		}

	}

	private static void allcountcontent(String string) {
		String[] filenames = new File(string).list();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(string + "//" + filename);
			}
			File f = new File(string + "//" + filename);
			if (f.isFile()) {
				Counter.allcount(string + "//" + filename);
			} else if (f.isDirectory()) {
				allcountcontent(string + "//" + filename);
			}
		}

	}

	private static void countContent(String string, int num) {
		String[] filenames = new File(string).list();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(string + "//" + filename);
			}
			File f = new File(string + "//" + filename);
			if (f.isFile()) {
				Counter.count(string + "//" + filename, num);
			} else if (f.isDirectory()) {
				countContent(string + "//" + filename, num);
			}
		}

	}

	private static void allFollows(String file, String cmd, int len) {
		String[] filenames = new File(file).list();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(file + "//" + filename);
			}
			File f = new File(file + "//" + filename);
			if (f.isFile()) {
				Counter.follows(file + "//" + filename, cmd, len);
			} else if (f.isDirectory()) {
				allFollows(file + "//" + filename, cmd, len);
			}
		}
	}

	private static void allPerceeds(String file, String cmd, int len) {
		String[] filenames = new File(file).list();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(file + "//" + filename);
			}
			File f = new File(file + "//" + filename);
			if (f.isFile()) {
				Counter.perceeds(file + "//" + filename, cmd, len);
			} else if (f.isDirectory()) {
				allPerceeds(file + "//" + filename, cmd, len);
			}
		}

	}

	private static void allEncodeLine(String file) {
		String[] filenames = new File(file).list();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(file + "//" + filename);
			}
			File f = new File(file + "//" + filename);
			if (f.isFile()) {
				result.add(LineEncoder.tostr(LineEncoder.encode(file + "//"
						+ filename)));
				// System.out.println(temp);
				// for (int i=0;i<temp.length;i++){
				// System.out.print(temp[i] +",");
				// }
				// System.out.println();
			} else if (f.isDirectory()) {
				allEncodeLine(file + "//" + filename);
			}
		}
//		result = new String[holder.size()];
//		for (int i = 0; i < holder.size(); i++) {
//			result[i] = holder.get(i);
//		}
//		System.out.println("i read " + result.length + " files");
//		return blah;
	}

	private static void llerapper(String file) {
		lens = new ArrayList<Integer>();
		lensEncodeLine(file);
		for (int i = 0; i < lens.size(); i++) {
			System.out.println(String.format("%5s,%5s", i, lens.get(i)));
		}

	}

	private static void lensEncodeLine(String file) {
		String[] filenames = new File(file).list();
		ArrayList<String> strholder = new ArrayList<String>();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(file + "//" + filename);
			}
			File f = new File(file + "//" + filename);
			if (f.isFile()) {
				strholder.add(LineEncoder.tostr(LineEncoder.encode(file + "//"
						+ filename)));
				int temp = strholder.get(strholder.size() - 1).length();
				while (temp >= lens.size()) {
					lens.add(0);
				}
				lens.set(temp, lens.get(temp) + 1);
			} else if (f.isDirectory()) {
				lensEncodeLine(file + "//" + filename);
			}
		}

	}

	private static void showMeTheMoney(String string) {

		String[] filenames = new File(string).list();
		ArrayList<String> project = new ArrayList<String>();
		for (String filename : filenames) {
			File f = new File(string + "//" + filename);
			if (f.isFile()) {
				project.add(string + "//" + filename);
			} else if (f.isDirectory()) {
				showMeTheMoney(string + "//" + filename);
			}
		}

		Counter.subs = new ArrayList<String>();
		Counter.fncs = new ArrayList<String>();
		Counter.dims = new ArrayList<String>();
		Counter.things = new ArrayList<String>();
		// Counter.fortemps = new ArrayList<String>();
		Counter.statics = new ArrayList<String>();
		Counter.consts = new ArrayList<String>();
		Counter.unkvars = new ArrayList<String>();
		Counter.unkthings = new ArrayList<String>();
		Counter.unknowns = new ArrayList<String>();

		String[][] ins = new String[project.size()][];
		String[][] incopy = new String[project.size()][];
		for (int i = 0; i < project.size(); i++) {
			ins[i] = Counter.flaten(project.get(i));
			incopy[i] = new String[ins[i].length];
			for (int j = 0; j < ins[i].length; j++) {
				incopy[i][j] = new String(ins[i][j]);
			}
			ins[i] = Counter.encode1(ins[i]);
		}
		for (int i = 0; i < project.size(); i++) {
			// System.out.println(project.get(i));
			System.out
					.println(Counter.tostr(Counter.encode2(ins[i], incopy[i])));
		}

	}

	public static void countsubstring(String[] tocount, String[] files) {
		int[] count = new int[tocount.length];
		int[] depth = new int[tocount.length];
		for (int i = 0; i < depth.length; i++) {
			depth[i] = 0;
			count[i] = 0;
		}
		for (int i = 0; i < files.length; i++) {
			for (int j = 0; j < files[i].length(); j++) {
				for (int k = 0; k < depth.length; k++) {
					if (files[i].charAt(j) == tocount[k].charAt(depth[k])) {
						depth[k]++;
						if (depth[k] >= tocount[k].length()) {
							count[k]++;
							depth[k] = 0;
						}
					} else {
						depth[k] = 0;
					}
				}
			}
		}
		for (int i = 0; i < count.length; i++) {
			System.out.println(String.format("%5s,%5s", tocount[i], count[i]));
		}
	}

	private static void countEncoded(String string) {
		String[] filenames = new File(string).list();
		ArrayList<String> project = new ArrayList<String>();
		for (String filename : filenames) {
			File f = new File(string + "//" + filename);
			if (f.isFile()) {
				project.add(string + "//" + filename);
			} else if (f.isDirectory()) {
				countEncoded(string + "//" + filename);
			}
		}

		Counter.subs = new ArrayList<String>();
		Counter.fncs = new ArrayList<String>();
		Counter.dims = new ArrayList<String>();
		Counter.things = new ArrayList<String>();
		// Counter.fortemps = new ArrayList<String>();
		Counter.statics = new ArrayList<String>();
		Counter.consts = new ArrayList<String>();
		Counter.unkvars = new ArrayList<String>();
		Counter.unkthings = new ArrayList<String>();
		Counter.unknowns = new ArrayList<String>();

		String[][] ins = new String[project.size()][];
		String[][] incopy = new String[project.size()][];
		for (int i = 0; i < project.size(); i++) {
			ins[i] = Counter.flaten(project.get(i));
			incopy[i] = new String[ins[i].length];
			for (int j = 0; j < ins[i].length; j++) {
				incopy[i][j] = new String(ins[i][j]);
			}
			ins[i] = Counter.encode1(ins[i]);
		}
		for (int i = 0; i < project.size(); i++) {
			System.out.println(project.get(i));
			ins[i] = Counter.encode2(ins[i], incopy[i]);
			Counter.countEncoded(ins[i]);
		}
	}

	private static int[][] cmpall() {
		int[][] scores = new int[result.size()][result.size()];
		for (int i = 0; i < result.size(); i++) {
			for (int j = 0; j < result.size(); j++) {
				if (i != j) {
					scores[i][j] = Cmp.cmp(result.get(i), result.get(i));
				} else {
					scores[i][j] =result.get(i).length() * 2 - 1;
				}
				System.out.println(result.size() * i + j + " of "
						+ result.size() * result.size());
			}
		}
		try {
			// Create file
			FileWriter fstream = new FileWriter("out.txt");
			BufferedWriter out = new BufferedWriter(fstream);
			for (int i = 0; i < scores.length; i++) {
				for (int j = 0; j < scores.length; j++) {
					out.write(String.format("%5s", scores[i][j]));
				}
				out.write("/n");
			}
			// Close the output stream
			out.close();
		} catch (Exception e) {// Catch exception if any
			System.err.println("Error: " + e.getMessage());
		}
		return scores;
	}
	
	public static int countEvents(String in){
		if (counts.containsKey(in)){
			return counts.get(in);
		}
		int c=0;
		for (int i=0;i<result.size();i++){
			for (int j=0;j<result.get(i).length() - in.length();j++){
				if (result.get(i).substring(j,j+in.length()).equals(in)){
					c++;
				}
			}
		}
		counts.put(in, c);
		return c;
	}
}
