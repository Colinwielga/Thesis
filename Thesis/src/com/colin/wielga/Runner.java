package com.colin.wielga;

import java.awt.List;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Scanner;

public class Runner {

	private static final String DIVIDER = "..........";
	public static boolean debug = false;
	public static ArrayList<Integer> lens = new ArrayList<Integer>();
	public static ArrayList<String> result = new ArrayList<String>();
	public static ArrayList<String> cheaters = new ArrayList<String>();
	public static HashMap<String, Integer> counts = new HashMap<String, Integer>();
	public static Runtime runTime;

	// for multi threading
	public static int[][] open;
	public static final int TOWRITE = 2;
	public static final int DONE = 1;
	public static final int WAITING = 0;
	public static final int WORKING = 3;
	public static double[][] mat;
	private static boolean canWrite = true;
	public static ArrayList<Gather> gathers = new ArrayList<Gather>();
	public static ArrayList<String> nameOrig= new ArrayList<String>();
	public static ArrayList<String> namePlag = new ArrayList<String>();

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		System.out.println("loading encoding system");
		LineEncoder.load("lineencoding1");

		// System.out.println("qc3 " + Cmp.qc3("ascaaasdradsf","vasdfasdfasd"));
		// System.out.println("qc " +
		// Cmp.qc_rapper("ascaaasdradsf","vasdfasdfasd"));
		startWorking("superPooper.txt", "superPooper.csv");
		keepWorking("superPooper.txt", "superPooper.csv");
		// System.out.println("done");
		// System.out.println("starting");
		// Cmp.qc_rapper("kllk",
		// "aaaaakmmmmkmkkkbkddqadtfbgifbpqaraaaaaaaagfbfbmksaoddaqacdaetcgefbmksaoddaqacdaetcgefbmksaoddaqacdaetcgefbddqaqacaaaaaaaaaaaaaaaettqadtfbddqaqacaaaaaaaaaaaaaaaettqadtfbddqaqacaaaaaaaaaaaaaaaettqadtfbaaf");
		// System.out.println("done");

		// Counter.load("C:\\Users\\Colin\\Documents\\School\\Thesis 2\\encoding1.txt");
		//
		//

		// runTime = Runtime.getRuntime();
		//
		// allEncodeLine("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
		// System.out.println("encoded all " + result.size() + " originals");
		// cheatersEncodeLine("C://Users//Colin//Documents//School//Thesis 2//VBCode//plaigarized");
		// System.out.println("encoded all " + cheaters.size() +
		// " plaigarized");
		// Alingment.loadvaluemat();
		// Alingment.localAl("acdcadsa", "dacdea");
		// doit("run3.txt");

		// System.out.println(Cmp.fastCmp("aabcabd", "baacada"));
		// s.nextInt();
		// System.out.println(java.lang.Runtime.getRuntime().maxMemory());
		//
		// char[] letters = {'a','b','c','d','e','f','g','h'};
		// //,'g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z'
		// for (int i=0;i<2000;i= i+2){
		// Long l = SpeedTest.test(SpeedTest.FASTCMP, i, letters, 1);
		// // System.out.println();
		// // if (i % 10 ==0){
		// System.out.println(l);// + " " + runTime.freeMemory()
		// // }
		//
		// }

		// System.out.println("" +
		// Cmp.fastCmp("aabacabdabsdbadsbbasbbdaabc","aabdsfbadbsbasbfbabsbsdbcbcabc"));

		// String[] temp =
		// LineEncoder.encode("VBcode\\original\\VBProjectsFall03\\aahumann\\frmPick.frm");
		// for (int i =0;i<temp.length;i++){
		// System.out.println(temp[i]);
		// }

		// allEncodeLine("C://Users//Colin//Documents//School//Thesis 2//VBCode//original");
		// cmpall();
		// allPerceeds("C://Users//Colin//Documents//School//Thesis 2//VBCode//original","i",1);
		// Counter.printSet();

		// String[] temp =
		// Counter.encode(Counter.flaten("C://Users//Colin//Documents//School//Thesis 2//VBCode//original//VBProjectsFall03//aahumann//frmPick.frm"));
		// for (int i=0;i<temp.length;i++){
		// System.out.print(temp[i] +",");
		// }

		// System.out.println("" +
		// Cmp.qc_rapper("acasdqaseasdqqafeadfawadasdveacasxczxcvdasdfeasdveghyagasdfvcrascvxawcefradvdacxvascvcxavzxcvawasdaqsasdfadasdfasdwdedasasdadce",
		// "eadasdffasqasdasfrdedwsadsfascasdcadsfaedcvascqacasdvaeacvadsvasdveadfeadsfasdfasdsdasdfbhbhbfasdfedssqaswdesdassdwwwsdasdwfedasesfwaeqdswwasadcada"));

		// countEncoded("C://Users//Colin//Documents//School//Thesis 2//VBCode//original//VBProjectsFall03//CsciStudent");

	}

	public void getin() {
		while (true) {
			System.out.println("ready for input");
			Scanner s = new Scanner(System.in);
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
		// result = new String[holder.size()];
		// for (int i = 0; i < holder.size(); i++) {
		// result[i] = holder.get(i);
		// }
		// System.out.println("i read " + result.length + " files");
		// return blah;
	}

	private static void cheatersEncodeLine(String file) {
		String[] filenames = new File(file).list();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(file + "//" + filename);
			}
			File f = new File(file + "//" + filename);
			if (f.isFile()) {
				cheaters.add(LineEncoder.tostr(LineEncoder.encode(file + "//"
						+ filename)));
				// System.out.println(temp);
				// for (int i=0;i<temp.length;i++){
				// System.out.print(temp[i] +",");
				// }
				// System.out.println();
			} else if (f.isDirectory()) {
				cheatersEncodeLine(file + "//" + filename);
			}
		}
		// result = new String[holder.size()];
		// for (int i = 0; i < holder.size(); i++) {
		// result[i] = holder.get(i);
		// }
		// System.out.println("i read " + result.length + " files");
		// return blah;
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
					scores[i][j] = result.get(i).length() * 2 - 1;
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

	private static void doit(String fileName) {
		FileWriter fstream;
		int temp;
		try {
			fstream = new FileWriter(fileName);
			BufferedWriter out = new BufferedWriter(fstream);
			for (int i = 0; i < result.size(); i++) {
				for (int j = 0; j < cheaters.size(); j++) {
					temp = Cmp.qc_rapper(result.get(i), cheaters.get(j));
					out.write(temp + ";");
					System.out.println(temp + "");
				}
				out.newLine();
			}
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static int countEvents(String in) {
		if (counts.containsKey(in)) {
			return counts.get(in);
		}
		int c = 0;
		for (int i = 0; i < result.size(); i++) {
			for (int j = 0; j < result.get(i).length() - in.length(); j++) {
				if (result.get(i).substring(j, j + in.length()).equals(in)) {
					c++;
				}
			}
		}
		counts.put(in, c);
		return c;
	}

	public static void startWorking(String meta, String data) {
		// get list of the files
		ArrayList resultsToWrite = startWorkingHelper(
				"C://Users//Colin//Documents//School//Thesis 2//VBCode//original",
				new ArrayList());
		resultsToWrite.addAll(startWorkingHelper(
				"C://Users//Colin//Documents//School//Thesis 2//VBCode//plaigarized//orig",
				new ArrayList()));
		ArrayList cheaterToWrite = startWorkingHelper(
				"C://Users//Colin//Documents//School//Thesis 2//VBCode//plaigarized//plag",
				new ArrayList());
		BufferedWriter bufferWritter;
		// write it to file
		try {
			bufferWritter = new BufferedWriter(new FileWriter(meta));
			// write the results
			for (int i = 0; i < resultsToWrite.size(); i++) {
				bufferWritter.write(resultsToWrite.get(i) + "");
				bufferWritter.newLine();
			}
			// write the DIVIDER
			bufferWritter.write(DIVIDER);
			bufferWritter.newLine();
			// write the cheaters
			for (int i = 0; i < cheaterToWrite.size(); i++) {
				bufferWritter.write(cheaterToWrite.get(i) + "");
				bufferWritter.newLine();
			}
			bufferWritter.close();

			// now create the data file
			bufferWritter = new BufferedWriter(new FileWriter(data, true));
			bufferWritter.write("");
			bufferWritter.close();

		} catch (IOException e) {
			e.printStackTrace();
			int chrash = 1 / 0;
		}
	}

	private static ArrayList startWorkingHelper(String file, ArrayList toReturn) {
		String[] filenames = new File(file).list();
		for (String filename : filenames) {
			if (debug) {
				System.out.println(file + "//" + filename);
			}
			File f = new File(file + "//" + filename);
			if (f.isFile()) {
				toReturn.add(file + "//" + filename);
			} else if (f.isDirectory()) {
				toReturn = startWorkingHelper(file + "//" + filename, toReturn);
			}
		}
		return toReturn;
	}

	public static void keepWorking(String meta, String data) {
		try {
			String read;

			// first we are going to load all the files
			BufferedReader in = new BufferedReader(new FileReader(meta));
			boolean go = true;

			System.out.println("encoding files");

			while ((read = in.readLine()) != null && go) {
				if (!DIVIDER.equals(read)) {
					// System.out.println(read);
					
					nameOrig.add(read);
					
					result.add(LineEncoder.tostr(LineEncoder.encode(read)));
				//	System.out.println(Analysis.getAdressEnd(read));
				} else {
					go = false;
				}
			}
			
			//System.out.println(DIVIDER);

			while ((read = in.readLine()) != null) {
				// System.out.println(read);
				namePlag.add(read);
				cheaters.add(LineEncoder.tostr(LineEncoder.encode(read)));
				//System.out.println(Analysis.getAdressEnd(read));
			}
			
			//int a=1;
			//while (1==a){
				//do nothing
			//}

			// find where you are
			in = new BufferedReader(new FileReader(data));
			int resultsCount = 0;
			int cheatersCount = 0;
			while ((read = in.readLine()) != null) {
				resultsCount = count(',', read);
				cheatersCount++;
				// System.out.println(cheatersCount);
			}

			if (resultsCount > 0) {
				resultsCount--;
			}

			System.out.println("calculating where to start");

			// System.out.println("got here");

			// resultsCount =1162;
			// cheatersCount = 24;
			open = new int[cheaters.size()][result.size()];
			mat = new double[cheaters.size()][result.size()];

			for (int i = 0; i < open.length; i++) {
				for (int j = 0; j < open[i].length; j++) {
					if (i < cheatersCount) {
						open[i][j] = DONE;
					} else if (i == cheatersCount && j < resultsCount) {
						open[i][j] = DONE;
					} else {
						open[i][j] = WAITING;
					}
				}
			}

			// System.out.println("we are at "+ resultsCount+ "/"+ result.size()
			// +" result");
			// System.out.println("we are at "+ cheatersCount+ "/"+
			// cheaters.size() +" cheaters");

			System.out.println("starting");

			// start comparing
			gathers.add(new Gather("postOp", data));
			gathers.add(new Gather("yo", data));
			gathers.add(new Gather("yim", data));
			gathers.add(new Gather("yuck", data));

			// for(int i=cheatersCount;i<cheaters.size();i++){
			// for (int j = resultsCount;j<result.size();j++){
			//
			// System.out.println("we are at "+ j + "/"+ result.size()
			// +" result "+ result.get(j));
			// System.out.println("we are at "+ i + "/"+ cheaters.size()
			// +" cheaters "+cheaters.get(i));
			//
			//
			// //cmp
			// String toright = Cmp.qc3(cheaters.get(i), result.get(j))+",";
			//
			// //save your work
			// BufferedWriter bufferWritter = new BufferedWriter(new
			// FileWriter(data,true));
			// bufferWritter.write(toright);
			// bufferWritter.close();
			//
			// }
			// // new line
			// BufferedWriter bufferWritter = new BufferedWriter(new
			// FileWriter(data,true));
			// bufferWritter.newLine();
			// bufferWritter.close();
			// //this is ugly
			// resultsCount = 0;
			// }
			// if we finish update the status file;

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			int chrash = 1 / 0;
		}

	}

	private static int count(char c, String read) {
		int result = 0;
		for (int i = 0; i < read.length(); i++) {
			if (read.charAt(i) == c) {
				result++;
			}
		}
		return result;
	}

	public static void TryWrite(String data) {
		if (canWrite) {
			canWrite = false;
			boolean go = true;
			for (int i = 0; go && i < open.length; i++) {
				for (int j = 0; go && j < open[i].length; j++) {
					if (open[i][j] == WORKING || open[i][j] == WAITING) {
						go = false;
					} else if (open[i][j] == TOWRITE) {
						BufferedWriter bufferWritter;
						try {
							bufferWritter = new BufferedWriter(new FileWriter(
									data, true));
							bufferWritter.write(mat[i][j] + ",");
							bufferWritter.close();

							if (j == open[i].length - 1) {
								bufferWritter = new BufferedWriter(
										new FileWriter(data, true));
								bufferWritter.newLine();
								bufferWritter.close();
							}
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
				}
			}
			canWrite  = true;
		}
	}
	
	public static void writeAll(double[][] mat2 ,String data){
		String nextWrite;
		try {
			BufferedWriter bufferWritter = new BufferedWriter(new FileWriter(data, true));

			for (int i=0;i<mat2.length;i++){
				nextWrite = "";
				for (int j =0;j<mat2[i].length;j++){
					nextWrite= nextWrite + mat2[i][j] +",";
				}
				bufferWritter.write(nextWrite);
				bufferWritter.newLine();
			}
			bufferWritter.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
}
