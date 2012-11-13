package com.colin.wielga;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map;
import java.util.Scanner;
import java.util.Map.Entry;
import java.util.Set;

public class Counter {
	//TODO these could also load from a file...
	private static final String DIMTO = "_v";
	private static final String FNCTO = "_f";
	private static final String SUBTO = "_s";
	private static final String TNGTO = "_t";
	private static final String UNKTO = "_u";
	static HashMap<String, Integer> c = new HashMap<String, Integer>();
	static ArrayList<String[]> fllws = new ArrayList<String[]>();

	static ArrayList<String> subs = new ArrayList<String>();
	static ArrayList<String> fncs = new ArrayList<String>();
	static ArrayList<String> dims = new ArrayList<String>();
	static ArrayList<String> things = new ArrayList<String>();
	// static ArrayList<String> fortemps = new ArrayList<String>();
	static ArrayList<String> statics = new ArrayList<String>();
	static ArrayList<String> consts = new ArrayList<String>();
	static ArrayList<String> unkvars = new ArrayList<String>();
	static ArrayList<String> unkthings = new ArrayList<String>();
	static ArrayList<String> unknowns = new ArrayList<String>();
	static HashMap<String, String> encoding = new HashMap<String, String>();

	public static void count(String f, int num) {
		// try {
		if (Runner.debug) {
			System.out.println("counting file " + f);
		}
		// Open the file that is the first
		// command line parameter
		File file = new File(f);
		FileInputStream fstream;
		try {
			fstream = new FileInputStream(file);
			// Get the object of DataInputStream
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;
			String[] s = new String[num];
			// Read File Line By Line
			while ((strLine = br.readLine()) != null) {
				// get the first element of the string
				String[] sp = devide(strLine);
				strLine = strLine.trim();
				int j = 0;
				for (int i = 0; i < num; i++) {
					s[i] = "";
					while (j < sp.length && sp[j].equals("")) {
						j++;
					}
					if (j < sp.length) {
						s[i] = sp[j];
					}
					j++;
				}
				add(sum(s));
			}

			// Close the input stream
			in.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// } catch (Exception e) {// Catch exception if any
		// System.err.println("Error: " + e.getMessage());
		// }
	}

	public static void allcount(String f) {
		// try {
		if (Runner.debug) {
			System.out.println("counting all file " + f);
		}
		// Open the file that is the first
		// command line parameter
		File file = new File(f);
		FileInputStream fstream;
		try {
			fstream = new FileInputStream(file);
			// Get the object of DataInputStream
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;
			// Read File Line By Line
			while ((strLine = br.readLine()) != null) {
				// get the first element of the string
				String[] sp = devide(strLine);
				for (int i = 0; i < sp.length; i++) {
					if (!sp[i].equals("") && !sp[i].equals(" ")) {
						add(sp[i]);
					}
				}
			}

			// Close the input stream
			in.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// } catch (Exception e) {// Catch exception if any
		// System.err.println("Error: " + e.getMessage());
		// }
	}

	public static void load(String f) {
		if (Runner.debug) {
			System.out.println("loading encoding " + f);
		}
		File file = new File(f);
		FileInputStream fstream;
		boolean onkey = true;
		String key = null;
		try {
			fstream = new FileInputStream(file);
			// Get the object of DataInputStream
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;
			// Read File Line By Line
			while ((strLine = br.readLine()) != null) {
				if (strLine.charAt(0) != '\'') {// a line that starts with ' is
												// a comment
					if (onkey) {
						key = strLine;
					} else {
						encoding.put(key, strLine);
						//if (Runner.debug) {
						//	System.out.println(key + " " + strLine);
						//}
					}
					onkey = !onkey;
				}

			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static String[] flaten(String f) {
		// TODO insert : between lines if last char was not " _" throw out " _"
		if (Runner.debug) {
			System.out.println("flatening file " + f);
		}
		// Open the file that is the first
		// command line parameter
		ArrayList<String> holder = new ArrayList<String>();
		File file = new File(f);
		FileInputStream fstream;
		try {
			fstream = new FileInputStream(file);
			// Get the object of DataInputStream
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;
			// Read File Line By Line
			while ((strLine = br.readLine()) != null) {
				// get the first element of the string
				String[] sp = devide(strLine);
				for (int i = 0; i < sp.length; i++) {
					if (!sp[i].equals("") && !sp[i].equals(" ")) {
						holder.add(sp[i]);
					}
				}
			}
			in.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		String[] result = new String[holder.size()];
		holder.toArray(result);
		return result;
	}

	public static void follows(String f, String s, int len) {
		String[] in = flaten(f);
		for (int i = 1; i < in.length; i++) {
			if (in[i - 1].equals(s)) {
				String[] toadd = new String[len];
				for (int j = 0; j < len; j++) {
					toadd[j] = in[i + j];
				}
				boolean go = true;
				for (int j = 0; go && j < fllws.size(); j++) {
					boolean nogo = true;
					for (int k = 0; k < len; k++) {
						if (i + j >= in.length) {
							toadd[j] = "";
						} else {
							toadd[j] = in[i + j];
						}
					}
					if (nogo) {
						go = false;
					}
				}
				if (go) {
					fllws.add(toadd);
				}
			}
		}
	}

	public static void perceeds(String f, String s, int len) {
		String[] in = flaten(f);
		for (int i = 1; i < in.length; i++) {
			if (in[i].equals(s)) {
				String[] toadd = new String[len];
				for (int j = 1; j < len + 1; j++) {
					if (i - j < 0) {
						toadd[j - 1] = "";
					} else {
						toadd[j - 1] = in[i - j];
					}
				}
				boolean go = true;
				for (int j = 0; go && j < fllws.size(); j++) {
					boolean nogo = true;
					for (int k = 0; k < len; k++) {
						if (!fllws.get(j)[k].equals(toadd[k])) {
							nogo = false;
						}
					}
					if (nogo) {
						go = false;
					}
				}
				if (go) {
					fllws.add(toadd);
				}
			}
		}
	}

	public static void printSet() {
		for (int j = 0; j < fllws.size(); j++) {
			for (int i = 0; i < fllws.get(j).length; i++) {
				System.out.print(fllws.get(j)[i] + " ");
			}
			System.out.println("");
		}
	}

	static String[] encode1(String[] in) {
		boolean stilldim = false;
		for (int i = 0; i < in.length; i++) {

			for (int j = 0; j < subs.size(); j++) {
				if (in[i].equals(subs.get(j))) {
					in[i] = "_";
					stilldim = false;
				}
			}
			for (int j = 0; j < fncs.size(); j++) {
				if (in[i].equals(fncs.get(j))) {
					in[i] = "_";
					stilldim = false;
				}
			}
			for (int j = 0; j < dims.size(); j++) {
				if (in[i].equals(dims.get(j))) {
					in[i] = "_";
					stilldim = false;
				}
			}
			for (int j = 0; j < things.size(); j++) {
				if (in[i].equals(things.get(j))) {
					in[i] = "_";
					stilldim = false;
				}
			}

			if (encoding.containsKey(in[i])) {
				if (!equalsin(in[i], ")", "(", "To", "As", "Boolean", "Byte",
						"Collection", "Currency", "Double", "Double", "Error",
						"Integer", "Long", "Object", "Single", "String",
						"User-Defined", "Variant")) {
					stilldim = false;
				}
				in[i] = encoding.get(in[i]);
				if (in[i].equals(encoding.get("Dim"))) {
					stilldim = true;
				}
			}

			if (isStr(in[i])) {
				in[i] = "_";
			} else if (isNum(in[i])) {
				in[i] = "_";
			}

			else if (in[i].charAt(0) == '#'
					&& isNum((String) in[i].subSequence(1, in[i].length()))) {
				in[i] = "_";
				stilldim = false;
				// System.out.println("found a file number: "+ in[i]);
			} else if (in[i].charAt(0) == ':'
					&& isHex((String) in[i].subSequence(1, in[i].length()))) {
				in[i] = "_";
				stilldim = false;
			}// what are these things?

			if (in[i].charAt(0) != '_') {
				if (i >= 2) {
					if (in[i - 2].equals("_.")) {
						if (equalsin(in[i - 1], 
								encoding.get("DriveListBox"),
								encoding.get("DirListBox"),
								encoding.get("FileListBox"),
								encoding.get("VScrollBar"),
								encoding.get("Shape"), 
								encoding.get("OLE"),
								encoding.get("ComboBox"),
								encoding.get("CommandButton"),
								encoding.get("Label"),
								encoding.get("CheckBox"),
								encoding.get("Image"), 
								encoding.get("Line"),
								encoding.get("PictureBox"),
								encoding.get("OptionButton"),
								encoding.get("Form"), 
								encoding.get("TextBox"),
								encoding.get("Timer"), 
								encoding.get("Menu"),
								encoding.get("Frame"), 
								encoding.get("ListBox"),
								encoding.get("HScrollBar"))) {
							things.add(in[i]);
							in[i] = "_";
							stilldim = false;
						}
					}
				}

				// i dont think we need fortemps...
				// if (i >= 1) {
				// if (in[i - 1].equals("_For")) {
				// if (i >= 2) {
				// if (in[i - 2].equals("_End")) {
				// // do nothing
				// } else {
				// System.out.println("found a for " +in[i]);
				// fortemps.add(in[1]);
				// in[i] = "_";
				// }
				// } else {
				// System.out.println("found a for " +in[i]);
				// fortemps.add(in[1]);
				// in[i] = "_";
				// }
				// }
				// }

				// TODO get the data type (maybe) vb is semi weakly typed so...
				if (i >= 1 && in[i].charAt(0) != '_') {
					if (in[i - 1].equals(encoding.get("Sub"))) {
						if (i >= 2) {
							if (in[i - 2].equals(encoding.get("End"))) {
								// do nothing
							} else {
								subs.add(in[1]);
								in[i] = "_";
								stilldim = false;
							}
						} else {
							subs.add(in[1]);
							in[i] = "_";
							stilldim = false;
						}
					}
					if (in[i - 1].equals(encoding.get("Sub"))) {
						subs.add(in[1]);
						in[i] = "_";
						stilldim = false;
					}
					if (in[i - 1].equals(encoding.get("Fnc"))) {
						if (i >= 2) {
							if (in[i - 2].equals(encoding.get("End"))) {
								// do nothing
							} else {
								fncs.add(in[1]);
								in[i] = "_";
								stilldim = false;
							}
						} else {
							fncs.add(in[1]);
							in[i] = "_";
							stilldim = false;
						}
					}
					if (in[i - 1].equals(encoding.get("Dim"))) {
						// TODO do i need to catch WithEvents
						dims.add(in[i]);
						in[i] = "_";
					}
					if (in[i - 1].equals(encoding.get("Const"))) {
						// TODO do i need to catch WithEvents
						// System.out.println("found a const " +in[i]);
						dims.add(in[i]);
						in[i] = "_";
					}
					if (in[i - 1].equals(encoding.get("Static"))) {
						// TODO we never find any of these...
						dims.add(in[i]);
						// System.out.println("found a static " +in[i]);
						in[i] = "_";
					}
					if (in[i - 1].equals(encoding.get("Pub"))) {
						// TODO i am not sure we need this .. but it does happen
						// in C://Users//Colin//Documents//School//Thesis
						// 2//VBCode//original//VBProjectsFall03//aahumann//Module1.bas
						// System.out.println("found a const " +in[i]);
						dims.add(in[i]);
						in[i] = "_i";
					}
				}

				if (i + 2 < in.length && in[i].charAt(0) != '_') {
					if (in[i + 1].equals("As")) {
						// this is somewhat questionable
						dims.add(in[i]);
					}
					if (in[i + 1].equals("(")) {
						int k = 2;
						while (!in[i + k].equals(")") && i + k + 1 < in.length) {
							k++;
						}
						k++;
						if (in[i + k].equals("As")) {
							dims.add(in[i]);
						}
					}
				}

				if (stilldim && in[i].charAt(0) != '_') {
					// System.out.println("found a dim with still dim "+in[i]);
					dims.add(in[i]);
					in[i] = "_";
				}

				if (i + 1 < in.length && in[i].charAt(0) != '_') {
					if (in[i + 1].equals("=")) {
						// its prolly a var
						unkvars.add(in[i]);
						// System.out.println("found unkvar " + in[i]);
					} else if (in[i + 1].equals("_.") || in[i + 1].equals(".")) {
						// its prolly a thing
						unkthings.add(in[i]);
						// System.out.println("found unkvar " + in[i]);
					} else {
						unknowns.add(in[i]);
					}
				} else {
					unknowns.add(in[i]);
				}
			}

		}

		return in;

	}

	private static boolean isHex(String s) {
		boolean result = true;
		for (int i = 0; result && i < s.length(); i++) {
			if (!(s.charAt(i) == '0' || s.charAt(i) == '1'
					|| s.charAt(i) == '2' || s.charAt(i) == '3'
					|| s.charAt(i) == '4' || s.charAt(i) == '5'
					|| s.charAt(i) == '6' || s.charAt(i) == '7'
					|| s.charAt(i) == '8' || s.charAt(i) == '9'
					|| s.charAt(i) == '0' || s.charAt(i) == 'A'
					|| s.charAt(i) == 'B' || s.charAt(i) == 'C'
					|| s.charAt(i) == 'D' || s.charAt(i) == 'E' || s.charAt(i) == 'F')) {
				result = false;
			}
		}
		return result;
	}

	public static String[] encode2(String[] in, String[] ogIn) {
		Scanner s = new Scanner(System.in);
		for (int i = 0; i < in.length; i++) {

			// TODO give vales to unknowns...
			// check for all the types of things...
			for (int j = 0; j < subs.size(); j++) {
				if (in[i].equals(subs.get(j))) {
					in[i] = SUBTO;
				}
			}
			for (int j = 0; j < fncs.size(); j++) {
				if (in[i].equals(fncs.get(j))) {
					in[i] = FNCTO;
				}
			}
			for (int j = 0; j < dims.size(); j++) {
				if (in[i].equals(dims.get(j))) {
					in[i] = DIMTO;
				}
			}
			for (int j = 0; j < things.size(); j++) {
				if (in[i].equals(things.get(j))) {
					in[i] = TNGTO;
				}
			}
			// for (int j = 0; j < fortemps.size(); j++) {
			// if (in[i].equals(fortemps.get(j))) {
			// in[i] = "_";
			// }
			// }
			for (int j = 0; j < unkvars.size(); j++) {
				if (in[i].equals(unkvars.get(j))) {
					in[i] = DIMTO;
				}
			}
			for (int j = 0; j < unkthings.size(); j++) {
				if (in[i].equals(unkthings.get(j))) {
					in[i] = TNGTO;
				}
			}

			for (int j = 0; j < unknowns.size(); j++) {
				if (in[i].equals(unknowns.get(j))) {
					in[i] = UNKTO;
				}
			}

			// if (in[i].charAt(0) != '_') {
			// for (int j = -2; j <= 2; j++) {
			// if (i + j < 0 || i + j >= in.length) {
			// System.out.print(",");
			// } else {
			// System.out.print(ogIn[i + j] + ",");
			// }
			// }
			// System.out.println();
			// s.next();
			// }
		}
		return in;
	}

	public static boolean equalsin(String b, String... a) {
		boolean result = false;
		for (int i = 0; i < a.length; i++) {
			if (a[i].equals(b)) {
				result = true;
			}
		}
		return result;
	}

	public static void countEncoded(String[] in) {
		for (int i = 0; i < in.length; i++) {
			add(in[i]);
		}
	}

	private static boolean isNum(String string) {
		try {
			NumberFormat format = NumberFormat.getInstance(Locale.US);
			Number number = format.parse(string);
			return true;
		} catch (Exception e) {
			// nop
		}
		if (string.charAt(0) == '&') {
			return true;
		}
		return false;
	}

	private static boolean isStr(String string) {
		if (string.charAt(0) == '"'
				&& string.charAt(string.length() - 1) == '"') {
			return true;
		}
		return false;
	}

	private static String[] devide(String string) {
		// split over " " and "," and "."
		// dont split inside " "
		// dont return empty strings
		// dont return things that start with '
		ArrayList<String> holder = new ArrayList<String>();
		string = string.trim();
		boolean quotes = false;
		boolean number = true;
		int lastcut = 0;
		int at = 0;
		while (at < string.length()) {
			if (string.charAt(at) == '"') {
				if (quotes) {
					at++;
					holder.add(string.substring(lastcut, at));
					number = true;
					while (at < string.length()
							&& (string.charAt(at) == ' ' || string.charAt(at) == ',')) {
						at++;
					}
					lastcut = at;
				} else {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
						number = false;
						lastcut = at;
					}
					at++;
				}
				quotes = !quotes;
			} else if (!quotes) {
				if (!(string.charAt(at) == '-' || string.charAt(at) == '0'
						|| string.charAt(at) == '1' || string.charAt(at) == '2'
						|| string.charAt(at) == '3' || string.charAt(at) == '4'
						|| string.charAt(at) == '5' || string.charAt(at) == '6'
						|| string.charAt(at) == '7' || string.charAt(at) == '8'
						|| string.charAt(at) == '9' || string.charAt(at) == '.')) {
					number = false;
				}
				if (string.charAt(at) == '\'') {
					String[] result = new String[holder.size()];
					holder.toArray(result);
					return result;
				}
				if (string.charAt(at) == ' ') {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
						number = true;
					}
					while (at < string.length()
							&& (string.charAt(at) == ' '
									|| string.charAt(at) == ','
									|| string.charAt(at) == '.' || string
									.charAt(at) == ';')) {
						at++;
					}
					lastcut = at;
				} else if (string.charAt(at) == ';') {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
						number = true;
					}
					while (at < string.length()
							&& (string.charAt(at) == ' '
									|| string.charAt(at) == ','
									|| string.charAt(at) == '.' || string
									.charAt(at) == ';')) {
						at++;
					}
					lastcut = at;
				} else if (string.charAt(at) == '=') {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
					}
					holder.add("=");
					number = true;
					at++;
					lastcut = at;
				} else if (string.charAt(at) == '(') {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
					}
					holder.add("(");
					number = true;
					at++;
					lastcut = at;
				} else if (string.charAt(at) == ')') {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
					}
					holder.add(")");
					number = true;
					at++;
					lastcut = at;
				} else if (string.charAt(at) == ',') {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
						number = true;
					}
					while (at < string.length()
							&& (string.charAt(at) == ' '
									|| string.charAt(at) == ','
									|| string.charAt(at) == '.' || string
									.charAt(at) == ';')) {// TODO these are
															// questionable when
															// would you see
															// this kinda shit
															// -> , , ..,., ,.
						at++;
					}
					lastcut = at;
				} else if (string.charAt(at) == '.' && number == false) {
					if (at != lastcut) {
						holder.add(string.substring(lastcut, at));
						holder.add(".");
						number = true;
					}
					while (at < string.length()
							&& (string.charAt(at) == ' '
									|| string.charAt(at) == ',' || string
									.charAt(at) == '.')) {
						at++;
					}
					lastcut = at;
				} else {
					at++;
				}
			} else {
				at++;
			}
		}
		String temp = "";
		for (int i = 0; i < at - lastcut; i++) {
			temp = temp + " ";
		}
		if (!string.substring(lastcut, at).equals(temp)) {
			holder.add(string.substring(lastcut, at));
		}

		// this wanted fix the Else: and PATH& problem but it catches
		// &H00004040& so we have a ugly catch for that
		for (int i = 0; i < holder.size(); i++) {
			if (holder.get(i).length() > 1) {
				if (holder.get(i).charAt(0) == '*'
						|| holder.get(i).charAt(0) == '-'
						|| holder.get(i).charAt(0) == '+'
						|| holder.get(i).charAt(0) == '/'
						|| (holder.get(i).charAt(0) == '&' && holder.get(i)
								.charAt(1) != 'H')
						|| holder.get(i).charAt(0) == '^'
						|| holder.get(i).charAt(0) == '.'
						|| (holder.get(i).charAt(0) == ':' && !isHex(holder
								.get(i).substring(1, holder.get(i).length())))) {
					//System.out.print("chopped " + holder.get(i));
					holder.add(i, holder.get(i).substring(0, 1));
					holder.set(
							i + 1,
							holder.get(i + 1).substring(1,
									holder.get(i + 1).length()));
					//System.out.println(" into " + holder.get(i) + " and "
					//		+ holder.get(i + 1));
				} else if (holder.get(i).charAt(0) == '*'
						|| holder.get(i).charAt(holder.get(i).length() - 1) == '-'
						|| holder.get(i).charAt(holder.get(i).length() - 1) == '+'
						|| holder.get(i).charAt(holder.get(i).length() - 1) == '/'
						|| (holder.get(i).charAt(holder.get(i).length() - 1) == '&' && holder
								.get(i).charAt(1) != 'H')
						|| holder.get(i).charAt(holder.get(i).length() - 1) == '^'
						|| holder.get(i).charAt(holder.get(i).length() - 1) == '.'
						|| holder.get(i).charAt(holder.get(i).length() - 1) == ':') {
					//System.out.print("chopped " + holder.get(i));
					holder.add(
							i + 1,
							holder.get(i).substring(holder.get(i).length() - 1,
									holder.get(i).length()));
					holder.set(
							i,
							holder.get(i).substring(0,
									holder.get(i).length() - 1));
					//System.out.println(" into " + holder.get(i) + " and "
					//		+ holder.get(i + 1));
				}

			}
		}

		for (int i = 0; i < holder.size();) {
			if (holder.get(i).equals("")) {
				holder.remove(i);
			} else {
				i++;
			}
		}

		String[] result = new String[holder.size()];
		holder.toArray(result);
		return result;
	}

	private static String sum(String[] s) {
		String result = "";
		for (int i = 0; i < s.length; i++) {
			if (!result.equals("")) {
				if (!s[i].equals("") && !s[i].equals(" ")) {
					result = result + "," + s[i];
				}
			} else {
				result = s[i];
			}
		}
		return result;

	}

	private static void add(String key) {
		if (!key.startsWith("'")) {
			if (Runner.debug) {
				System.out.println("added key " + key);
			}
			if (c.containsKey(key)) {
				c.put(key, c.get(key) + 1);
			} else {
				c.put(key, 1);
			}
		}
	}
 
	public static void getTotals() {
		Iterator it = c.entrySet().iterator();
		ArrayList<String> keys = new ArrayList<String>();
		ArrayList<Integer> values = new ArrayList<Integer>();
		while (it.hasNext()) {
			Map.Entry pairs = (Map.Entry) it.next();
			int i = 0;
			while (i < values.size() && values.get(i) < (int) pairs.getValue()) {
				i++;
			}
			keys.add(i, (String) pairs.getKey());
			values.add(i, (Integer) pairs.getValue());
			it.remove(); // avoids a ConcurrentModificationException
		}
		for (int i = 0; i < keys.size(); i++) {
			System.out.println(keys.get(i) + "  " + values.get(i));
		}
	}
	
	public static String tostr(String[] in){
		String result = "";
		for(int i=0;i<in.length;i++){
			result = result + in[i].substring(1, in[i].length());
		}
		return result;
	}

	public static void getChecked() {
		Iterator it = c.entrySet().iterator();
		ArrayList<String> keys = new ArrayList<String>();
		ArrayList<Integer> values = new ArrayList<Integer>();
		while (it.hasNext()) {
			Map.Entry pairs = (Map.Entry) it.next();
			int i = 0;
			while (i < values.size() && values.get(i) < (int) pairs.getValue()) {
				i++;
			}
			keys.add(i, (String) pairs.getKey());
			values.add(i, (Integer) pairs.getValue());
			it.remove(); // avoids a ConcurrentModificationException
		}
		for (int i = 0; i < keys.size(); i++) {
			if (keys.get(i).charAt(0) != '_') {
				System.out.println(keys.get(i) + "  " + values.get(i));
			}
		}

	}
}
