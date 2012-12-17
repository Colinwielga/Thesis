package com.colin.wielga;

import java.util.ArrayList;
import java.util.Arrays;

public class Cmp {

	public static void cmpProject(String[] a, String[] b) {

	}

	public static int qc_rapper(String a, String b) {
		boolean[][] big = new boolean[a.length()][b.length()];
		// generate the big matrix
		for (int i = 0; i < a.length(); i++) {
			for (int j = 0; j < b.length(); j++) {
				big[i][j] = a.charAt(i) == b.charAt(j);
			}
		}
		// printbig(big);
		System.out.println("");
		return (qc(big));
	}

	public static int fastCmp(String a, String b) {
		Straight winner = new Straight();
		boolean[][] big = new boolean[a.length()][b.length()];
		// generate the big matrix
		for (int i = 0; i < a.length(); i++) {
			for (int j = 0; j < b.length(); j++) {
				big[i][j] = a.charAt(i) == b.charAt(j);
			}
		}
		//printbig(big);
		
		ArrayList<Straight> straights = findStraights(big); // i wrote this
															// withn my nose
		//System.out.println(""+straights.size());
//		for (int i =0 ; i<straights.size();i++){
//			straights.get(i).print();
//		}
		
		Straight[] temp;
		Straight target = null;
		if (straights.size() != 0) {
			target = straights.get(0);
		}
		while (target != null) {
			temp = cross(straights, target);
			Straight[] straightstemp = new Straight[straights.size()];
			straights.toArray(straightstemp);
			//System.out.println("looking at a new Straight");
			BestResult br = best(null, straightstemp, temp, straights.get(0));
			// update winner
			//System.out.println("and the winner is... (HOV)");
			//br.winner.print();
			winner = new Straight(winner, br.winner);
			// remove all the elements of br.dissallows
			//System.out.println(""+br.dissallows.length);
			for (int i = 0; i < br.dissallows.length; i++) {
				for (int j = straights.size()-1; j > -1; j--)
					// we need to count down so removing does not cause problems
					if (straights.get(j).eqls(br.dissallows[i])) {
						straights.remove(j);
						//System.out.println("i removed somthing!");br.dissallows[i].print();
					}
			}
			// get the next target
			target = null;
			if (straights.size() != 0) {
				target = straights.get(0);
			}
		}
		return winner.value();
	}

	private static Straight[] cross(ArrayList<Straight> straights,
			Straight straight) {
		ArrayList<Straight> result = new ArrayList<Straight>();
		for (int i = 0; i < straights.size(); i++) {
			if (!straight.coexist(straights.get(i))){
				result.add(straights.get(i));
			}
		}
		Straight[] ret = new Straight[result.size()];
		result.toArray(ret);
		return ret;
	}

	private static Straight[] cross(Straight[] straights, Straight straight) {
		ArrayList<Straight> result = new ArrayList<Straight>();
		for (int i = 0; i < straights.length; i++) {
			if (!straight.coexist(straights[i]))
				result.add(straights[i]);
		}
		Straight[] ret = new Straight[result.size()];
		result.toArray(ret);
		return ret;
	}

	public static ArrayList<Straight> findStraights(boolean[][] in) {
		// find all the Straights sorted by lenght
		ArrayList<Straight> straights = new ArrayList<Straight>();
		for (int i = 0; i < in.length; i++) {
			for (int j = 0; j < in[i].length; j++) {
				if (in[i][j]) {
					int k = 1;
					boolean go = true;
					while (go) {
						int l = 0;
						boolean look = true;
						while (look && l < straights.size()) {
							if (scale(k) < straights.get(l).value) {
								l++;
							} else {
								look = false;
							}
						}
						straights.add(l, new Straight(i, j, i + k, j + k, k));
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
		return straights;
	}

	public static int qc(boolean[][] in) {
		// find all the Straights
		ArrayList<Straight> straights = findStraights(in);
		// pull out all the biggest one ...if two are tied and overlap take the
		// one we found first
		Straight currentStraight;
		ArrayList<Straight> result = new ArrayList<Straight>();
		while (straights.size() != 0) {
			currentStraight = straights.get(0);
			straights.remove(0);
			// ArrayList<Straight> lookingat = new ArrayList<Straight>();
			for (int i = 0; i < straights.size();) {
				if (!currentStraight.coexist(straights.get(i))) {
					straights.remove(i);
				}
			}
			result.add(currentStraight);
		}
		// add up what we got...
		int resultsum = 0;
		for (int i = 0; i < result.size(); i++) {
			resultsum = resultsum + result.get(i).value();
		}
		// print the winning matrix
		// boolean[][] toprint = new boolean[in.length][in[0].length];
		// for (int i = 0; i < toprint.length; i++) {
		// for (int j = 0; j < toprint[i].length; j++) {
		// toprint[i][j] = false;
		// }
		// }
		// for (int i = 0; i < result.size(); i++) {
		// for (int j = 0; j < result.get(i).len; j++) {
		// toprint[result.get(i).ystart + j][result.get(i).xstart + j] = true;
		// }
		// }
		// printbig(toprint);
		return resultsum;
	}

	// TODO i am writing an algorithm that check to see if we know a striaght
	// larger than any set of striaghts in its cross
	public static BestResult best(StraightCoexistMat coexist, Straight[] all,
			Straight[] s, Straight biggie) {
//		System.out.println("biggie is ");
//		biggie.print();
//		System.out.println("is s empty " + s.length);
		int tobeat = biggie.value;
		// first we update the co-exist mat
		if (coexist == null) {
			coexist = new StraightCoexistMat(s.length);
		}
		int at = 0;
		ArrayList<Integer> shrink = new ArrayList<Integer>();

		// TODO if i am only ever passing in super sets this maybe need to be
		// rewrote
		// rearange coexist.straights
		for (int l = 0; l < coexist.size(); l++) {
			for (int m = 0; m < s.length; m++) {
				if (s[m].holds(coexist.getStraight(l))) {
					// TODO should be if s[m].contains(coexists.getStraight(l));
					coexist.straights.set(at, s[m]);
					shrink.add(l);
					at++;
				}
			}
		}
		// now revalue coexist.mat
		for (int l = 0; l < shrink.size(); l++) {
			for (int m = 0; m < shrink.size(); m++) {
				coexist.set(1, m, coexist.get(shrink.get(l), shrink.get(m)));
			}
		}
		coexist.doneUpTo = shrink.size()-1;
		
		// now update coexists straights so they match what we want them to
		ArrayList<Straight> newstraights = new ArrayList<Straight>();
		for (int l = 0; l <shrink.size();l++){
			newstraights.add(coexist.getStraight(shrink.get(l)));
		}
		for (int l=0 ;l<s.length;l++){
			if (!shrink.contains(s[l])){
				newstraights.add(s[l]);
			}
		}
		coexist.straights = newstraights;
		
		// now fill in the rest of
		for (int i = 0; i < coexist.size(); i++) {
			// coexist.add(new ArrayList<Boolean>());
			for (int j = 0; j < coexist.size(); j++) {
				// TODO this only really need to be half filled but that might
				// mess up our ref? just make a big boolean[][]?
				if (coexist.doneUpTo < i && coexist.doneUpTo < j) {
					coexist.set(i, j, s[i].coexist(s[j]));
				}
			}
		}
		// add a check to see if any single straight can beat the one we have
		// (?)
		// do some of dat shit
		ArrayList<Straight> tie = new ArrayList<Straight>();
		tie.add(biggie);
		//System.out.println("coexist size is "+coexist.size());
		for (int i = 0; i < coexist.size(); i++) {
			for (int j = i + 1; j < coexist.size(); j++) {
				if (coexist.get(i, j)) {
					coexist.set(i, j, false);
					coexist.insert(new Straight(coexist.getStraight(i), coexist
							.getStraight(j)));
					if (coexist.value(i) + coexist.value(j) > tobeat) {
						tobeat = coexist.value(i) + coexist.value(j);
						tie = new ArrayList<Straight>();
						tie.add(coexist.getStraight(coexist.size()-1));
					}
					if (coexist.value(i) + coexist.value(j) == tobeat) {
						tie.add(coexist.getStraight(coexist.size()-1));
						// we will deal if the ties if we need to at the end
					}
					// add the new values
					for (int k = 0; k < coexist.size(); k++) {
						coexist.set(k, coexist.size() - 1, coexist.get(k, i)
								&& coexist.get(k, j));
						if (coexist.get(k, i) && coexist.get(k, j)) {
							coexist.set(k, i, false);
							coexist.set(k, j, false);
						}
						coexist.set(coexist.size() - 1, k, coexist.get(i, k)
								&& coexist.get(j, k));
						if (coexist.get(i, k) && coexist.get(j, k)) {
							coexist.set(i, k, false);
							coexist.set(j, k, false);
						}
					}
				}
			}
		}
		// resolve the ties
		for (int i = 0; i < tie.size(); i++) {
			// itterate over the ties, if their cross is not contain in one
			// the early crosses we
			// pass back something to look at ?
			// we look at it ? <- this one
			boolean isnew = false;
			Straight[] c1 = cross(all, tie.get(i));
			// first check to see if its contained in s (the cross we are
			// looking at)
			boolean isIn;
			for (int j = 0; !isnew && j < c1.length; j++) {
				isIn = false;
				for (int k = 0; !isIn && k < s.length; k++) {
					if (s[k].eqls(c1[j])) {
						isIn = true;
					}
				}
				if (!isIn) {
					//System.out.println(c1[j] + "in s");
					isnew = true;
				}
			}
			// TODO we might be able to save time here by checking len of c1
			// and now check to see if it is contained in any of the tie we have
			// already looked at
			
			// now that i think about it, i dont think i need this at all
//			for (int j = 0; !isnew && j < i; j++) {
//				Straight[] c2 = cross(all, tie.get(j));
//				for (int k = 0; !isnew && k < c2.length; k++) {
//					isIn = false;
//					for (int l = 0; !isIn && l < c1.length; l++) {
//						if (c2[k].eqls(c1[l])) {
//							isIn = true;
//						}
//					}
//					if (!isIn) {
//						System.out.println(c1[j] + "in c2");
//						isnew = true;
//					}
//				}
//			}
			if (isnew) {
				// we want to try tie.get(i) on the intersection of its cross
				// and biggies cross
				// does this break anything?
				Straight[] yolo = intersect(c1, s); // i believe this should
													// just be the same as c1
				BestResult feedback = best(coexist, all, yolo, tie.get(i));
				if (feedback.foundone) {
					return new BestResult(feedback.winner, yolo);
				} else {
					// i think we should never have an else
					System.out.println("this should not happen");
				}
			}
		}
		// if none of the ties add to the cross they all have the same affect
		// choose any of them
		return new BestResult(tie.get(0), s);
	}

	private static Straight[] intersect(Straight[] s1, Straight[] s2) {
		ArrayList<Straight> holder = new ArrayList<Straight>();
		for (int i = 0; i < s1.length; i++) {
			if (!holder.contains(s1[i])) {
				holder.add(s1[i]);
			}
		}
		for (int i = 0; i < s2.length; i++) {
			if (!holder.contains(s2[i])) {
				holder.add(s2[i]);
			}
		}
		Straight[] result = new Straight[holder.size()];
		holder.toArray(result);
		return result;
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

	public static int scale(int count) {
		// TODO Auto-generated method stub
		if (count == 0) {
			return 0;
		}
		return (2 * count - 1);
	}
}
