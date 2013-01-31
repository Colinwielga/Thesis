package com.colin.wielga;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class Alingment {

	static int[][] valuemat;
	static List<Character> axis = new ArrayList<Character>();
	static BufferedReader br;
	private static String sCurrentLine;
	private static String[] nextstep;

	public static void loadvaluemat() {
		try {
			br = new BufferedReader(new FileReader("LAinfo"));
			int count = 0;
			while (!(sCurrentLine = br.readLine()).equals("=====")) {

				if (LineEncoder.hm.containsKey(sCurrentLine)) {
					axis.add(LineEncoder.hm.get(sCurrentLine).charAt(0));
					System.out.println(LineEncoder.hm.get(sCurrentLine).charAt(
							0));
				} else {
					axis.add(' ');// never happens
					System.out.println(" ");
				}
			}
			axis.add('-');
			count = 0;
			valuemat = new int[axis.size()][axis.size()];
			while ((sCurrentLine = br.readLine()) != null) {
				nextstep = sCurrentLine.split(",");
				for (int i = 0; i < nextstep.length; i++) {
					valuemat[count][i] = Integer.parseInt(nextstep[i]);
				}
				count++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		// make sure it worked:
//		for (int i = 0; i < axis.size(); i++) {
//			System.out.println(i + " " + axis.get(i));
//		}
//		for (int i = 0; i < valuemat.length; i++) {
//			for (int j = 0; j < valuemat[i].length; j++) {
//				System.out.print(valuemat[i][j] + ",");
//			}
//			System.out.println();
//		}

	}

	public static int globalAl(String a, String b) {
		int[][] pathmat = new int[a.length() + 1][b.length() + 1];
		pathmat[0][0] = 0;
		for (int i = 0; i < a.length(); i++) {
			pathmat[i + 1][0] = valuemat[getAxisPos(a.charAt(i))][getAxisPos('-')]
					+ pathmat[i][0];
		}
		for (int i = 0; i < b.length(); i++) {
			pathmat[0][i + 1] = valuemat[getAxisPos('-')][getAxisPos(b
					.charAt(i))] + pathmat[0][i];
		}

		for (int i = 0; i < b.length(); i++) {
			for (int j = 0; j < a.length(); j++) {
				// System.out.println(j + " "+i);
				pathmat[j + 1][i + 1] = max(
						pathmat[j][i]
								+ valuemat[getAxisPos(a.charAt(j))][getAxisPos(b
										.charAt(i))],
						pathmat[j + 1][i]
								+ valuemat[getAxisPos(a.charAt(j))][getAxisPos('-')],
						pathmat[j][i + 1]
								+ valuemat[getAxisPos('-')][getAxisPos(b
										.charAt(i))]);
			}
		}

		// print pathmat
//		for (int i = 0; i < pathmat.length; i++) {
//			for (int j = 0; j < pathmat[i].length; j++) {
//				System.out.print(pathmat[i][j] + ",");
//			}
//			System.out.println();
//		}

		// TODO it might be nice to be able to see what is the winning result
		return pathmat[a.length()][b.length()];

	}

	public static int localAl(String a, String b) {
		int[][] pathmat = new int[a.length() + 1][b.length() + 1];
		pathmat[0][0] = 0;
		for (int i = 0; i < a.length(); i++) {
			pathmat[i + 1][0] = 0;
		}
		for (int i = 0; i < b.length(); i++) {
			pathmat[0][i + 1] = 0;
		}

		int currentMax =0;
		for (int i = 0; i < b.length(); i++) {
			for (int j = 0; j < a.length(); j++) {
				// System.out.println(j + " "+i);
				int temp = max(
						pathmat[j][i]
								+ valuemat[getAxisPos(a.charAt(j))][getAxisPos(b
										.charAt(i))],
						pathmat[j + 1][i]
								+ valuemat[getAxisPos(a.charAt(j))][getAxisPos('-')],
						pathmat[j][i + 1]
								+ valuemat[getAxisPos('-')][getAxisPos(b
										.charAt(i))], 0);
				pathmat[j + 1][i + 1] = temp;
				if (temp>currentMax){
					currentMax = temp;
				}
			}
		}

		// print pathmat
//		for (int i = 0; i < pathmat.length; i++) {
//			for (int j = 0; j < pathmat[i].length; j++) {
//				System.out.print(pathmat[i][j] + ",");
//			}
//			System.out.println();
//		}

		// TODO it might be nice to be able to see what is the winning result
		return currentMax;

	}

	private static int max(int i, int j, int k, int l) {
		if (l <= i) {
			if (i <= j) {
				if (j <= k) {
					return k;
				} else {
					return j;
				}
			} else {
				if (i >= k) {
					return i;
				} else {
					return k;
				}
			}
		} else {
			if (l <= j) {
				if (j <= k) {
					return k;
				} else {
					return j;
				}
			} else {
				if (l >= k) {
					return l;
				} else {
					return k;
				}
			}
		}
	}

	private static int getAxisPos(char c) {
		int ret = 0;
		while (axis.get(ret) != c) {
			ret++;
		}
		return ret;
	}

	private static int max(int i, int j, int k) {
		if (i <= j) {
			if (j <= k) {
				return k;
			} else {
				return j;
			}
		} else {
			if (i >= k) {
				return i;
			} else {
				return k;
			}
		}
	}

}
