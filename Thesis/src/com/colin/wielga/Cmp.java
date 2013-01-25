package com.colin.wielga;

import java.util.ArrayList;
import java.util.Arrays;

public class Cmp {

	public static StraightCoexistArray coexist = new StraightCoexistArray();

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
		System.out.println("make big");
		// printbig(big);
		// System.out.println("");
		return (qc(big, a, b));
	}

	// public static int qc3(String a, String b){
	// boolean[][] big = new boolean[a.length()][b.length()];
	// for (int i=0;i<big.length;i++){
	// for (int j = 0;j<big[i].length;j++){
	// big[i][j] = a.charAt(i) == b.charAt(j);
	// }
	// }
	//
	// }

	public static double qc3(String a, String b) {
		boolean[][] big = new boolean[a.length()][b.length()];
		boolean OneInTheOven = false;
		int startx = 0;
		int starty = 0;
		ArrayList<Straight> straights = new ArrayList<Straight>();
		// first we need to gill big, and find the straights
		for (int i = 0; i < a.length() + b.length() - 1; i++) {
			for (int j = 0; j < Math.min(Math.min((i + 1), (a.length())),
					Math.min((b.length()), (a.length() + b.length() - 1 - i))); j++) {
				// System.out.println(Math.max(a.length() -1 - i + j, j) + " , "
				// + (j + Math.max(0,-a.length() + 1+i)));
				big[Math.max(a.length() - 1 - i + j, j)][(j + Math.max(0,
						-a.length() + 1 + i))] = a.charAt(Math.max(a.length()
						- 1 - i + j, j)) == b.charAt((j + Math.max(0,
						-a.length() + 1 + i)));

				if (OneInTheOven
						&& big[Math.max(a.length() - 1 - i + j, j)][(j + Math
								.max(0, -a.length() + 1 + i))]
						&& (a.charAt(Math.max(a.length() - 1 - i + j, j)) + "")
								.equals(LineEncoder.hm.get(LineEncoder.ENDSUB))) {
					// end our striaght including the current element
					straights = addSorted(
							straights,
							new Straight(
									startx,
									starty,
									(j + Math.max(0, -a.length() + 1 + i)) + 1,
									Math.max(a.length() - 1 - i + j, j) + 1,
									(j + Math.max(0, -a.length() + 1 + i) - startx) + 1));
					OneInTheOven = false;
				} else if (OneInTheOven
						&& big[Math.max(a.length() - 1 - i + j, j)][(j + Math
								.max(0, -a.length() + 1 + i))]
						&& (a.charAt(Math.max(a.length() - 1 - i + j, j)) + "")
								.equals(LineEncoder.hm
										.get(LineEncoder.STARTSUB))) {
					// end our straight not inculding the space we are currently
					// at, then start a new one
					straights = addSorted(
							straights,
							new Straight(
									startx,
									starty,
									(j + Math.max(0, -a.length() + 1 + i)),
									Math.max(a.length() - 1 - i + j, j),
									(j + Math.max(0, -a.length() + 1 + i) - startx)));
					startx = (j + Math.max(0, -a.length() + 1 + i));
					starty = Math.max(a.length() - 1 - i + j, j);
				} else if (OneInTheOven
						&& !big[Math.max(a.length() - 1 - i + j, j)][(j + Math
								.max(0, -a.length() + 1 + i))]) {
					// add a straight
					straights = addSorted(
							straights,
							new Straight(
									startx,
									starty,
									(j + Math.max(0, -a.length() + 1 + i)),
									Math.max(a.length() - 1 - i + j, j),
									(j + Math.max(0, -a.length() + 1 + i) - startx)));
					// System.out.println("made one " +
					// startx+" , "+starty+" , "+(j + Math.max(0,-a.length() +
					// 1+i))+" , "+Math.max(a.length() -1 - i + j, j)+ " , " +
					// ((j + Math.max(0,-a.length() + 1+i))-startx));
					OneInTheOven = false;
				} else if (!OneInTheOven
						&& big[Math.max(a.length() - 1 - i + j, j)][(j + Math
								.max(0, -a.length() + 1 + i))]) {
					// start a new straight
					OneInTheOven = true;
					startx = (j + Math.max(0, -a.length() + 1 + i));
					starty = Math.max(a.length() - 1 - i + j, j);
				}
			}
			// if we have one unfinished
			if (OneInTheOven) {
				// add it
				int temp = Math.min(Math.min((i + 1), (a.length())), Math.min(
						(b.length()), (a.length() + b.length() - 1 - i)));
				straights = addSorted(
						straights,
						new Straight(startx, starty, (temp + Math.max(0,
								-a.length() + 1 + i)), Math.max(a.length() - 1
								- i + temp, temp), (temp + Math.max(0,
								-a.length() + 1 + i))
								- startx));
				// System.out.println("end loop made one "+
				// startx+" , "+starty+" , "+(temp + Math.max(0,-a.length() +
				// 1+i))+" , "+ (a.length() - i - 1 +temp) +" , "+ ((temp +
				// Math.max(0,-a.length() + 1+i)) - startx));
				OneInTheOven = false;
			}
			// shoot we need to sort
		}

		// check the starting mat
		// boolean[][] toprint = new boolean[a.length()][b.length()];
		// for (int i = 0; i < toprint.length; i++) {
		// for (int j = 0; j < toprint[i].length; j++) {
		// toprint[i][j] = false;
		// }
		// }
		// for (int i = 0; i < straights.size(); i++) {
		// for (int j = 0; j < straights.get(i).len[0]; j++) {
		// if (toprint[straights.get(i).xstart[0] +
		// j][straights.get(i).ystart[0] + j]){
		// System.out.println("its already true... Straight finding is fu");
		// int temp = 1/0;
		// }
		// toprint[straights.get(i).xstart[0] + j][straights.get(i).ystart[0] +
		// j] = true;
		// }
		// }
		// for (int i =0 ;i < big.length;i++){
		// for (int j=0;j<big[i].length;j++){
		// if ((a.charAt(i) == b.charAt(j)) != toprint[i][j]){
		// System.out.println("shit is bad at "+i +" "+j);
		// }
		// }
		// }
		// printbig(toprint);

		Straight currentStraight;
		ArrayList<Straight> subtraction;
		ArrayList<Straight> result = new ArrayList<Straight>();
		while (straights.size() != 0) {

			currentStraight = straights.get(0);
			// System.out.println("added "+currentStraight.xstart[0] +" "+
			// currentStraight.ystart[0] +" "+ currentStraight.xend[0] +" "+
			// currentStraight.yend[0] +" "+ currentStraight.len[0]);
			straights.remove(0);
			// ArrayList<Straight> lookingat = new ArrayList<Straight>();
			for (int i = straights.size() - 1; i >= 0; i--) {
				if (!currentStraight.coexist(straights.get(i))) {
					subtraction = straightSub(straights.get(i), currentStraight);
					// System.out.println("removed "+straights.get(i).xstart[0]
					// +" "+ straights.get(i).ystart[0] +" "+
					// straights.get(i).xend[0] +" "+ straights.get(i).yend[0]
					// +" "+ straights.get(i).len[0]);
					straights.remove(i);
					for (int j = 0; j < subtraction.size(); j++) {
						// System.out.println("anding remainder "+subtraction.get(j).xstart[0]
						// +" "+ subtraction.get(j).ystart[0] +" "+
						// subtraction.get(j).xend[0] +" "+
						// subtraction.get(j).yend[0] +" "+
						// subtraction.get(j).len[0]);
						straights = addSorted(straights, subtraction.get(j));
					}
				}
			}
			result.add(currentStraight);

			// check the current matrix
			// toprint = new boolean[a.length()][b.length()];
			// for (int i = 0; i < toprint.length; i++) {
			// for (int j = 0; j < toprint[i].length; j++) {
			// toprint[i][j] = false;
			// }
			// }
			// for (int i = 0; i < result.size(); i++) {
			// for (int j = 0; j < result.get(i).len[0]; j++) {
			// toprint[result.get(i).xstart[0] + j][result.get(i).ystart[0] + j]
			// = true;
			// }
			// }
			//
			// for (int i = 0; i < toprint.length; i++) {
			// for (int j = 0; j < toprint[i].length; j++) {
			// if (toprint[i][j]){
			// for (int k=0;k<toprint.length;k++){
			// if (toprint[k][j] && k != i){
			// System.out.println("qc3 result sucks");
			// int crash = 1/0;
			// }
			// }
			// for (int k=0;k<toprint[i].length;k++){
			// if (toprint[i][k] && k != j){
			// System.out.println("qc3  result sucks");
			// int crash = 1/0;
			// }
			// }
			// }
			//
			// }
			// }

		}
		// System.out.print("found a solution");

		// add up what we got...
		double resultsum = 0;
		int totalLen = 0;
		for (int i = 0; i < result.size(); i++) {
			// System.out.println(result.get(i).xstart[0] +" "+
			// result.get(i).ystart[0] +" "+ result.get(i).xend[0]+" "+
			// result.get(i).yend[0] );
			resultsum = resultsum + result.get(i).value();
			for (int j = 0; j < result.get(i).len.length; j++) {
				totalLen = totalLen + result.get(i).len[j];
			}
		}
		
		//for unmatched dudes
		//resultsum = resultsum + totalLen*2 - a.length() - b.length();
		

		// System.out.println();
		// // check the winning matrix
		// toprint = new boolean[a.length()][b.length()];
		// for (int i = 0; i < toprint.length; i++) {
		// for (int j = 0; j < toprint[i].length; j++) {
		// toprint[i][j] = false;
		// }
		// }
		// for (int i = 0; i < result.size(); i++) {
		// for (int j = 0; j < result.get(i).len[0]; j++) {
		// toprint[result.get(i).xstart[0] + j][result.get(i).ystart[0] + j] =
		// true;
		// }
		// }
		//
		//
		// for (int i = 0; i < toprint.length; i++) {
		// for (int j = 0; j < toprint[i].length; j++) {
		// if (toprint[i][j]){
		// for (int k=0;k<toprint.length;k++){
		// if (toprint[k][j] && k != i){
		// System.out.println("qc3 result sucks");
		// int crash = 1/0;
		// }
		// }
		// for (int k=0;k<toprint[i].length;k++){
		// if (toprint[i][k] && k != j){
		// System.out.println("qc3  result sucks");
		// int crash = 1/0;
		// }
		// }
		// }
		//
		// }
		// }
		// //printbig(toprint);
		// return

		return resultsum;
	}

	private static ArrayList<Straight> addSorted(ArrayList<Straight> straights,
			Straight straight) {

		// System.out.println("adding " + straight.len[0]);
		//
		// for (int i=0;i<straights.size();i++){
		// System.out.print(straights.get(i).len[0] + " , ");
		// }
		// System.out.println();

		int above = 0;
		int below = straights.size();
		int at;
		while (true) {
			at = (int) Math.floor((above + below) / 2);
			// System.out.println("stuck " + at +" " + above + " " + below);
			if (below <= 0) {
				straights.add(0, straight);
				// System.out.println("inserted at 0");

				// for (int i=0;i<straights.size()-1;i++){
				// if (straights.get(i).len[0] < straights.get(i+1).len[0]){
				// System.out.println("we are not sorted!");
				// int crash = 1/0;
				// }
				// }

				return straights;
			} else if (at >= straights.size()) {
				straights.add(straight);
				// System.out.println("inserted at "+ straights.size());

				// for (int i=0;i<straights.size()-1;i++){
				// if (straights.get(i).len[0] < straights.get(i+1).len[0]){
				// System.out.println("we are not sorted!");
				// int crash = 1/0;
				// }
				// }

				return straights;
			} else if (below == above) {
				straights.add(at, straight);
				// System.out.println("inserted at "+ at);

				// for (int i=0;i<straights.size()-1;i++){
				// if (straights.get(i).len[0] < straights.get(i+1).len[0]){
				// System.out.println("we are not sorted!");
				// int crash = 1/0;
				// }
				// }

				return straights;
			}
			if (straights.get(at).len[0] == straight.len[0]) {
				// cool we are done.
				straights.add(at, straight);
				// System.out.println("inserted at "+ at);

				// for (int i=0;i<straights.size()-1;i++){
				// if (straights.get(i).len[0] < straights.get(i+1).len[0]){
				// System.out.println("we are not sorted!");
				// int crash = 1/0;
				// }
				// }

				return straights;
			} else if (straights.get(at).len[0] > straight.len[0]) {
				above = at + 1;
			} else { // straights.get(at).len[0] > straight.len[0]
				below = at;
			}
		}

		// let's check to make sure we are sorted
	}

	private static ArrayList<Straight> straightSub(Straight s, Straight t) {
		// its s - t

		boolean[] sThing = new boolean[s.len[0]];
		for (int i = 0; i < sThing.length; i++) {
			if (t.xstart[0] <= i + s.xstart[0] && i + s.xstart[0] < t.xend[0]) {
				sThing[i] = false;
			} else if (t.ystart[0] <= i + s.ystart[0]
					&& i + s.ystart[0] < t.yend[0]) {
				sThing[i] = false;
			} else {
				sThing[i] = true;
			}
		}

		ArrayList<Straight> toReturn = new ArrayList<Straight>();
		boolean praggers = false;
		int start = 0;
		for (int i = 0; i < sThing.length; i++) {
			if (praggers && !sThing[i]) {
				toReturn.add(new Straight(s.ystart[0] + start, s.xstart[0]
						+ start, s.ystart[0] + i, s.xstart[0] + i, i - start));
				// System.out.println(" added 1 "+(s.xstart[0]+start) + " " +
				// (s.ystart[0]+start) +" "+(s.xstart[0]+i)+" "+(s.ystart[0]+i)
				// +" "+ (i-start) );
				praggers = false;
			} else if (!praggers && sThing[i]) {
				start = i;
				praggers = true;
			}
		}
		if (praggers) {
			toReturn.add(new Straight(s.ystart[0] + start, s.xstart[0] + start,
					s.ystart[0] + sThing.length, s.xstart[0] + sThing.length,
					sThing.length - start));
			// System.out.println(" added 2 "+(s.xstart[0]+start) + " " +
			// (s.ystart[0]+start)+" "+(s.xstart[0]+sThing.length)+" "+(s.ystart[0]+sThing.length)+" "+(sThing.length
			// - start));
		}

		return toReturn;
	}

	// this is a helper for qc2 which does notdoes not work
	public static ArrayList addStraight(ArrayList<Straight> winners, Straight s) {

		// System.out.println("thinking about adding a thing "+(i-len)+" "+(j-len)+" "+(i+1)+" "+(j+1));
		boolean innerGo = true;
		for (int k = 0; innerGo && k < winners.size(); k++) {
			innerGo = (winners.get(k).coexist(s) || winners.get(k).value < s.value);
			if (!innerGo) {
				winners.get(k).killed.add(s);
				System.out.println("blocked by " + winners.get(k).xstart[0]
						+ " " + winners.get(k).ystart[0] + " "
						+ winners.get(k).xend[0] + " " + winners.get(k).yend[0]
						+ " " + winners.get(k).coexist(s));
			}
		}
		// if nothing kills it remove all the ones smaller then it and then add
		// it
		if (innerGo) {
			for (int k = winners.size() - 1; k >= 0; k--) {
				if (!winners.get(k).coexist(s)
						&& winners.get(k).value < s.value) {
					// System.out.println("removed "+winners.get(k).xstart[0]+" "+
					// winners.get(k).ystart[0]+" "+winners.get(k).xend[0]+" "+winners.get(k).yend[0]
					// + " " + winners.get(k).coexist(s));
					for (int l = 0; l < winners.get(k).killed.size(); l++) {
						if (winners.get(k).killed.get(l).coexist(s)) {
							winners = addStraight(winners,
									winners.get(k).killed.get(l));
						}
					}
					s.killed.add(winners.get(k));
					winners.remove(k);
				}
			}
			winners.add(s);
			// System.out.println("added a thing, we now have " +
			// winners.size());
		}

		return winners;
	}

	// this does not work
	public static int qc2(String a, String b) {
		boolean[][] big = new boolean[a.length()][b.length()];
		ArrayList<Straight> winners = new ArrayList<Straight>();
		Straight s;
		// generate the big matrix
		for (int i = 0; i < a.length(); i++) {
			for (int j = 0; j < b.length(); j++) {
				if (a.charAt(i) == b.charAt(j)) {
					big[i][j] = true;
					// we found a set of straigh lets see how big it is
					boolean go = true;
					int len = 0;
					while (go) {
						// check to see if s is disallowed by a bigger straight
						s = new Straight(i - len, j - len, i + 1, j + 1,
								1 + len);
						winners = addStraight(winners, s);
						// if we have reached the start of a sub, stop.
						if ((a.charAt(i - len) + "").equals(LineEncoder.hm
								.get(LineEncoder.STARTSUB))) {
							go = false;
						}
						// move on to the next straight
						len++;
						// if we have gone outside the mat stop
						if (i - len < 0 || j - len < 0) {
							go = false;
							// if we have reached the end of the straight or ...
						} else if (!big[i - len][j - len]
								|| (a.charAt(i - len) + "")
										.equals(LineEncoder.hm
												.get(LineEncoder.ENDSUB))) {
							go = false;
						}

					}

				} else {
					big[i][j] = false;
				}

			}
		}
		int resultsum = 0;
		for (int i = 0; i < winners.size(); i++) {
			resultsum = resultsum + winners.get(i).value();
		}
		// System.out.println(resultsum);
		return resultsum;
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
		// printbig(big);

		ArrayList<Straight> straights = findStraights(big, a, b); // i wrote
																	// this
		// withn my nose
		// System.out.println(""+straights.size());
		// for (int i =0 ; i<straights.size();i++){
		// straights.get(i).print();
		// }

		Straight[] temp;
		Straight target = null;
		if (straights.size() != 0) {
			target = straights.get(0);
		}
		while (target != null) {
			temp = cross(straights, target);
			Straight[] straightstemp = new Straight[straights.size()];
			straights.toArray(straightstemp);
			// System.out.println("looking at a new Straight");
			coexist.reset();
			BestResult br = best(straightstemp, temp, straights.get(0));
			// update winner
			// System.out.println("and the winner is... (HOV)");
			// br.winner.print();
			winner = new Straight(winner, br.winner);
			// remove all the elements of br.dissallows
			// System.out.println(""+br.dissallows.length);
			for (int i = 0; i < br.dissallows.length; i++) {
				for (int j = straights.size() - 1; j > -1; j--)
					// we need to count down so removing does not cause problems
					if (straights.get(j).eqls(br.dissallows[i])) {
						straights.remove(j);
						// System.out.println("i removed somthing!");br.dissallows[i].print();
					}
			}
			// get the next target
			target = null;
			if (straights.size() != 0) {
				target = straights.get(0);
			}
		}
		// for (int i =0;i<a.length();i++){
		// for (int j=0;j<b.length();j++){
		// System.out.print(winner.pointIn(i,j));
		// }
		// System.out.println();
		// }
		return winner.value();
	}

	private static Straight[] cross(ArrayList<Straight> straights,
			Straight straight) {
		ArrayList<Straight> result = new ArrayList<Straight>();
		for (int i = 0; i < straights.size(); i++) {
			if (!straight.coexist(straights.get(i))) {
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

	public static ArrayList<Straight> findStraights(boolean[][] in, String a,
			String b) {
		// find all the Straights sorted by lenght
		ArrayList<Straight> straights = new ArrayList<Straight>();
		for (int i = 0; i < in.length; i++) {
			for (int j = 0; j < in[i].length; j++) {
				if (in[i][j]) {
					int k = 1;
					boolean go = true;
					while (go) {
						// System.out.println("outer while " +i +" "+j +" "+k);
						int l = 0;
						boolean look = true;
						while (look && l < straights.size()) {
							// System.out.println("inner while " + l);
							if (scale(k) < straights.get(l).value) {
								l++;
							} else {
								look = false;
							}
						}
						straights.add(l, new Straight(i, j, i + k, j + k, k));
						if (i + k < in.length && j + k < in[0].length) {
							if (in[i + k][j + k]
									&& !("" + a.charAt(i + k))
											.equals(LineEncoder.hm
													.get(LineEncoder.STARTSUB))) {
								if (("" + a.charAt(i + k))
										.equals(LineEncoder.hm
												.get(LineEncoder.ENDSUB))) {
									go = false;
								}
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

		// check the starting mat
		// boolean[][] toprint = new boolean[a.length()][b.length()];
		// for (int i = 0; i < toprint.length; i++) {
		// for (int j = 0; j < toprint[i].length; j++) {
		// toprint[i][j] = false;
		// }
		// }
		// for (int i = 0; i < straights.size(); i++) {
		// for (int j = 0; j < straights.get(i).len[0]; j++) {
		// toprint[straights.get(i).ystart[0] + j][straights.get(i).xstart[0] +
		// j] = true;
		// }
		// }
		// for (int i =0 ;i < in.length;i++){
		// for (int j=0;j < in[i].length;j++){
		// if ((a.charAt(i) == b.charAt(j)) != toprint[i][j]){
		// System.out.println("shit is bad at "+i +" "+j);
		// }
		// }
		// }

		return straights;
	}

	public static int qc(boolean[][] in, String a, String b) {
		// find all the Straights
		ArrayList<Straight> straights = findStraights(in, a, b);
		// System.out.println("found Straights");
		// pull out all the biggest one ...if two are tied and overlap take the
		// one we found first
		Straight currentStraight;
		ArrayList<Straight> result = new ArrayList<Straight>();
		while (straights.size() != 0) {
			// System.out.println(straights.size());
			currentStraight = straights.get(0);
			straights.remove(0);
			// ArrayList<Straight> lookingat = new ArrayList<Straight>();
			for (int i = 0; i < straights.size();) {
				if (!currentStraight.coexist(straights.get(i))) {
					straights.remove(i);
				} else {
					i++;
				}
			}
			result.add(currentStraight);
		}
		// System.out.print("found a solution");
		// add up what we got...
		int resultsum = 0;
		for (int i = 0; i < result.size(); i++) {
			resultsum = resultsum + result.get(i).value();
		}
		// System.out.println("added things up");
		// print the winning matrix
		boolean[][] toprint = new boolean[in.length][in[0].length];
		for (int i = 0; i < toprint.length; i++) {
			for (int j = 0; j < toprint[i].length; j++) {
				toprint[i][j] = false;
			}
		}
		// for (int i = 0; i < result.size(); i++) {
		// for (int j = 0; j < result.get(i).len[0]; j++) {
		// toprint[result.get(i).ystart[0] + j][result.get(i).xstart[0] + j] =
		// true;
		// }
		// }
		// printbig(toprint);

		for (int i = 0; i < toprint.length; i++) {
			for (int j = 0; j < toprint[i].length; j++) {
				if (toprint[i][j]) {
					for (int k = 0; k < toprint.length; k++) {
						if (toprint[k][j] && k != i) {
							System.out.println("qc result sucks");
							int crash = 1 / 0;
						}
					}
					for (int k = 0; k < toprint[i].length; k++) {
						if (toprint[i][k] && k != j) {
							System.out.println("qc result sucks");
							int crash = 1 / 0;
						}
					}
				}

			}
		}

		return resultsum;
	}

	// TODO i am writing an algorithm that check to see if we know a striaght
	// larger than any set of striaghts in its cross
	public static BestResult best(Straight[] all, Straight[] s, Straight biggie) {
		// System.out.println("biggie is ");
		// biggie.print();
		// System.out.println("is s empty " + s.length);
		int tobeat = biggie.value;
		// first we update the co-exist mat
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
		coexist.doneUpTo = shrink.size() - 1;

		// now update coexists straights so they match what we want them to
		ArrayList<Straight> newstraights = new ArrayList<Straight>();
		for (int l = 0; l < shrink.size(); l++) {
			newstraights.add(coexist.getStraight(shrink.get(l)));
		}
		for (int l = 0; l < s.length; l++) {
			if (!shrink.contains(s[l])) {
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
		// System.out.println("coexist size is "+coexist.size());
		for (int i = 0; i < coexist.size(); i++) {
			for (int j = i + 1; j < coexist.size(); j++) {
				if (coexist.get(i, j)) {
					coexist.set(i, j, false);
					coexist.insert(new Straight(coexist.getStraight(i), coexist
							.getStraight(j)));
					if (coexist.value(i) + coexist.value(j) > tobeat) {
						tobeat = coexist.value(i) + coexist.value(j);
						tie = new ArrayList<Straight>();
						tie.add(coexist.getStraight(coexist.size() - 1));
					}
					if (coexist.value(i) + coexist.value(j) == tobeat) {
						tie.add(coexist.getStraight(coexist.size() - 1));
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
					// System.out.println(c1[j] + "in s");
					isnew = true;
				}
			}
			// TODO we might be able to save time here by checking len of c1
			// and now check to see if it is contained in any of the tie we have
			// already looked at

			// now that i think about it, i dont think i need this at all
			// for (int j = 0; !isnew && j < i; j++) {
			// Straight[] c2 = cross(all, tie.get(j));
			// for (int k = 0; !isnew && k < c2.length; k++) {
			// isIn = false;
			// for (int l = 0; !isIn && l < c1.length; l++) {
			// if (c2[k].eqls(c1[l])) {
			// isIn = true;
			// }
			// }
			// if (!isIn) {
			// System.out.println(c1[j] + "in c2");
			// isnew = true;
			// }
			// }
			// }
			if (isnew) {
				// we want to try tie.get(i) on the intersection of its cross
				// and biggies cross
				// does this break anything?
				Straight[] yolo = intersect(c1, s); // i believe this should
													// just be the same as c1
				BestResult feedback = best(all, yolo, tie.get(i));
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

	private static double startwork(boolean[][] big, int depth, double currentvalue) {
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
		double tempvalue = currentvalue;
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
					double temp = startwork(b2, depth + 1, currentvalue);
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
				double temp = startwork(b2, depth + 1, currentvalue);
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

	public static double scale(int count) {
		// TODO Auto-generated method stub
		if (count == 0) {
			return 0;
		}
		// return (2 * count - 1);
		 return count * count;
		//return fac(count);
	}

	private static double fac(int count) {
		// TODO Auto-generated method stub
		if (count == 0 ){
			return 1;
		}
		return (count*fac(count-1));
	}
}
