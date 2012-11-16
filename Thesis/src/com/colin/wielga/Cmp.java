package com.colin.wielga;

import java.util.ArrayList;
import java.util.Arrays;

public class Cmp {

	public static void cmpProject(String[] a, String[] b) {

	}
	
	public static int qc_rapper(String a, String b){
		boolean[][] big = new boolean[a.length()][b.length()];
		// generate the big matrix
		for (int i = 0; i < a.length(); i++) {
			for (int j = 0; j < b.length(); j++) {
				big[i][j] = a.charAt(i) == b.charAt(j);
			}
		}
		//printbig(big);
		System.out.println("");
		return(qc(big));
	}
	
	public static int qc(boolean[][] in) {
		// find all the Straights
		ArrayList<Straight> straights = new ArrayList<Straight>();
		for (int i = 0; i < in.length; i++) {
			for (int j = 0; j < in[i].length; j++) {
				if (in[i][j]) {
					int k = 1;
					boolean go = true;
					while (go) {
						int l=0;
						boolean look = true;
						while (look && l<straights.size()){
							if (k<straights.get(l).len){
								l++;
							}else {
								look = false;
							}
						}
						straights.add(l,new Straight(i, j, i + k, j + k, k));
						if (i + k < in.length && j + k < in.length) {
							if (in[i + k][j + k]) {
								k++;
							} else {
								go = false;
							}
						} else {
							go = false;
						}
					}
				}
			}
		}
		// pull out all the biggest one ...if two are tied and overlap take the one we found first
		Straight currentStraight;
		ArrayList<Straight> result = new ArrayList<Straight>();
		while (straights.size() != 0) {
			currentStraight = straights.get(0);
			straights.remove(0);
			// ArrayList<Straight> lookingat = new ArrayList<Straight>();
			for (int i = 0; i < straights.size();) {
				if (currentStraight.xstart <= straights.get(i).xstart
						&& straights.get(i).xstart < currentStraight.xend) {
					straights.remove(i);
				} else if (currentStraight.xstart < straights.get(i).xend
						&& straights.get(i).xend <= currentStraight.xend) {
					straights.remove(i);
				} else if (currentStraight.ystart <= straights.get(i).ystart
						&& straights.get(i).ystart < currentStraight.yend) {
					straights.remove(i);
				} else if (currentStraight.ystart < straights.get(i).yend
						&& straights.get(i).yend <= currentStraight.yend) {
					straights.remove(i);
				} else {
					i++;
				}
			}
			result.add(currentStraight);
		}
		// add up what we got...
		int resultsum = 0;
		for (int i=0;i<result.size();i++){
			resultsum = resultsum + scale(result.get(i).len);
		}
		boolean[][] toprint = new boolean[in.length][in[0].length];
		for (int i=0;i<toprint.length;i++){
			for (int j=0;j<toprint[i].length;j++){
				toprint[i][j] = false;
			}
		}
		for (int i=0;i<result.size();i++){
			for (int j=0;j<result.get(i).len;j++){
				toprint[result.get(i).ystart + j][result.get(i).xstart + j] = true;
			}
		}
		//printbig(toprint);
		return resultsum;
	}

	public static int cmp(String a, String b) {
		boolean[][] big = new boolean[a.length()][b.length()];
		// generate the big matrix
		for (int i = 0; i < a.length(); i++) {
			for (int j = 0; j < b.length(); j++) {
				big[i][j] = a.charAt(i) == b.charAt(j);
			}
		}

		// we can split it up in to chunks that share no lines -- no we can't
		ArrayList<Integer> cuts = new ArrayList<Integer>();
		cuts.add(0);
		for (int i = 0; i < big.length; i++) {
			boolean split = true;
			for (int j = 0; split && j < big.length; j++) {
				if (big[i][j]) {
					// if (i>0 && j>0){
					// if (big[i-1][j-1]){
					// split = false;
					// }
					// }
					if (i < big.length - 1 && j < big.length - 1) {
						if (big[i - 1][j - 1]) {
							split = false;
						}
					}
				}
			}
			if (split) {
				cuts.add(i);
			}
		}
		cuts.add(big.length);
		// now we run startwork on all the chunks
		boolean[][] b2;
		ArrayList<Combo[]> pieces = new ArrayList<Combo[]>();
		for (int i = 1; i < cuts.size(); i++) {
			b2 = new boolean[cuts.get(i) - cuts.get(i - 1)][b.length()];
			for (int j = 0; j < cuts.get(i) - cuts.get(i - 1); j++) {
				for (int k = 0; k < b.length(); k++) {
					b2[j][k] = big[cuts.get(i - 1) + j][k];
				}
			}
			pieces.add(getPieces(b2));
			// but start work need to return all the possible met and thier
			// scores
			// we then need to find the best combo of the cuts
		}
		ArrayList<int[]> tolookat = new ArrayList<int[]>();
		int[] at = new int[pieces.size()];
		for (int i = 0; i < at.length; i++) {
			at[i] = 0; // pieces.get(i).length;
		}
		tolookat.add(at);
		boolean foundOne = false;
		int best = 0;
		int r = 0;
		ArrayList<int[]> nextlookat = new ArrayList<int[]>();
		while (!foundOne) {
			for (int i = 0; i < tolookat.size(); i++) {
				r = tryCombine(at, pieces);
				if (r != -1) {
					foundOne = true;
					if (r > best) {
						best = r;
					}
				}
			}
			if (!foundOne) {
				// if we have not found one we need to generate a new list of
				// thing to search
				for (int i = 0; i < tolookat.size(); i++) {
					for (int j = 0; j < tolookat.get(i).length; j++) {
						int[] totry = deepCopy(tolookat.get(i));
						totry[j] = totry[j] - 1;
						boolean add = true;
						for (int k = 0; add && k < nextlookat.size(); k++) {
							boolean allmatch = true;
							for (int l = 0; allmatch && l < totry.length; l++) {
								if (totry[l] != nextlookat.get(k)[l]) {
									allmatch = false;
								}
							}
							if (allmatch) {
								add = false;
							}
						}
						if (add) {
							nextlookat.add(totry);
						}
					}
				}
				tolookat = nextlookat;
			}
		}
		return best;

	}

	private static int[] deepCopy(int[] in) {
		int[] result = new int[in.length];
		for (int i = 0; i < in.length; i++) {
			result[i] = in[i];
		}
		return result;
	}

	private static int tryCombine(int[] at, ArrayList<Combo[]> pieces) {
		// we need to see if the current at can combine

		// combine them..
		int len = 0;
		for (int i = 0; i < pieces.size(); i++) {
			len = len + pieces.get(i)[0].mat[0].length;
		}
		boolean[][] combine = new boolean[pieces.get(0)[0].mat.length][len];
		int in = 0;
		for (int i = 0; i < pieces.size(); i++) {
			for (int j = 0; i < pieces.get(j)[at[i]].mat[0].length; j++) {
				for (int k = 0; k < pieces.get(j)[at[i]].mat.length; k++) {
					combine[k][in] = pieces.get(j)[at[i]].mat[k][j];
				}
				in++;
			}
		}

		// now check to see it they hold (have only a single one in each row)
		boolean foundone = false;
		for (int i = 0; i < combine.length; i++) {
			int count = 0;
			for (int j = 0; j < combine[i].length; j++) {
				if (combine[i][j]) {
					if (count == 1) {
						foundone = true;
					} else {
						count = 1;
					}
				}
			}
		}
		// if we found a solution return it
		if (foundone) {
			int sumScores = 0;
			for (int i = 0; i < at.length; i++) {
				sumScores = sumScores + pieces.get(i)[at[i]].score;
			}
			return sumScores;
		}// else{
			// if we did not find a solution we return -1
		return -1;
	}

	private static Combo[] getPieces(boolean[][] b2) {
		// Combop[] is sorted.
		// TODO Auto-generated method stub
		return null;
	}

	private static int startwork(boolean[][] big, int depth, int currentvalue) {
		// update current value if we can...
		if (depth > 1) {
			int lastlastat = -1;
			for (int i = 0; lastlastat == -1 && i < big.length; i++) {
				if (big[i][depth - 2]) {
					lastlastat = i;
				}
			}
			int lastat = -1;
			for (int i = 0; lastat == -1 && i < big.length; i++) {
				if (big[i][depth - 1]) {
					lastat = i;
				}
			}
			// we can only update if the last two don't line up
			if (lastat != lastlastat + 1 && lastat != -1 && lastlastat != -1) {
				int count = 0;
				boolean go = true;
				// count how many are in a row
				for (int i = depth - 2; go && i >= 0; i--) {
					if (lastlastat - (depth - 2) + i >= 0 && i >= 0) {
						if (big[lastlastat - (depth - 2) + i][i]) {
							count++;
						} else {
							go = false;
						}
					} else {
						go = false;
					}
				}
				// System.out.println("i added " + scale(count));
				currentvalue = scale(count) + currentvalue;
			}
		}

		// TODO think about also recording the best big..
		boolean[][] b2;
		boolean emptyrow = true;
		boolean needallzero = true;
		int tempvalue = currentvalue;
		if (depth < big[0].length) {
			// for all the options run a copy of this
			for (int i = 0; i < big.length; i++) {
				if (big[i][depth]) {
					emptyrow = false;
					b2 = deepCopy(big);
					// now that we have picked an point we 0 out everything else
					// in its collum
					for (int j = 0; j < big.length; j++) {
						b2[j][depth] = i == j;
					}
					// and everything else in its row
					boolean otherOptions = false;
					for (int j = depth; j < big[0].length; j++) {
						if (j > depth && b2[i][j]) {
							otherOptions = true;
						}
						b2[i][j] = depth == j;
					}
					if (!otherOptions) {
						needallzero = false;
					}
					int temp = startwork(b2, depth + 1, currentvalue);
					if (temp > tempvalue) {
						tempvalue = temp;
						// System.out.println("depth= " +depth +" i found " +
						// temp);
					}
					// if (depth == b2[0].length - 1) {
					// printbig(b2);
					// System.out.println(temp);
					// }

				}
			}
			// if we need to try the emmpty row
			if (emptyrow || needallzero) {
				b2 = deepCopy(big);
				for (int i = 0; i < big.length; i++) {
					b2[i][depth] = false;
				}
				int temp = startwork(b2, depth + 1, currentvalue);
				if (temp > tempvalue) {
					tempvalue = temp;
				}
				// if (depth == b2[0].length - 1) {
				// printbig(b2);
				// System.out.println(temp);
				// }
			}
			currentvalue = tempvalue;
		} else {
			// we are done but we need to evalute the last chunk
			int lastlastat = -1;
			for (int i = 0; lastlastat == -1 && i < big.length; i++) {
				if (big[i][depth - 1]) {
					lastlastat = i;
				}
			}
			int count = 0;
			boolean go = true;
			// count how many are in a row
			// System.out.println(depth + " " + big[0].length);
			if (lastlastat != -1) {
				for (int i = depth - 1; go && i >= 0; i--) {
					if ((lastlastat - (depth - 1) + i) >= 0 && i >= 0) {
						// System.out.println((lastlastat - (depth - 1) + i) +
						// " "
						// + i);
						if (big[lastlastat - (depth - 1) + i][i]) {
							count++;
						} else {
							go = false;
						}
					} else {
						go = false;
					}
				}
			}
			// System.out.println("i added " + scale(count));
			currentvalue = scale(count) + currentvalue;
		}
		// System.out.println(depth+" "+currentvalue);
		return currentvalue;
	}

	private static void printbig(boolean[][] big) {
		for (int i = 0; i < big.length; i++) {
			for (int j = 0; j < big[i].length; j++) {
				if (big[i][j]) {
					System.out.print(1);
				} else {
					System.out.print(0);
				}
			}
			System.out.println();
		}
	}

	private static boolean[][] deepCopy(boolean[][] original) {
		if (original == null) {
			return null;
		}

		final boolean[][] result = new boolean[original.length][];
		for (int i = 0; i < original.length; i++) {
			result[i] = Arrays.copyOf(original[i], original[i].length);
			// For Java versions prior to Java 6 use the next:
			// System.arraycopy(original[i], 0, result[i], 0,
			// original[i].length);
		}
		return result;
	}

	private static int scale(int count) {
		// TODO Auto-generated method stub
		if (count == 0) {
			return 0;
		}
		return (2 * count - 1);
	}
}
