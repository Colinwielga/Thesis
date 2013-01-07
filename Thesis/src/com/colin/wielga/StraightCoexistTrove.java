package com.colin.wielga;

import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import gnu.trove.list.TIntList;
import gnu.trove.list.array.TIntArrayList;

public class StraightCoexistTrove {
	private static final int STARTSIZE = 5000;
	private static final int GROWBY = 1000;
	public TIntList[] trove;
	ArrayList<Straight> straights = new ArrayList<Straight>();
	public int doneUpTo;
//	public StraightCoexistArray check = new StraightCoexistArray();

	public StraightCoexistTrove() {
		reset();
	}

	public void reset() {
		trove = new TIntArrayList[STARTSIZE];
		for (int i = 0; i < STARTSIZE; i++) {
			trove[i] = new TIntArrayList();
		}
		straights = new ArrayList<Straight>();
		doneUpTo = 0;
	}

	public boolean get(int y, int x) {
		if (y < x) {
			int c = x;
			x = y;
			y = c;
		}
//		if (y == 101 && x ==1) {
//			System.out.println("getting " + x + " " + y);
//			for (int i = 0; i < trove[x].size(); i++) {
//				System.out.print(trove[x].get(i) + " ");
//			}
//			System.out.println();
//		}

		int upper = trove[x].size() - 1;
		int lower = 0;
		int at;
		// int counter=0;
		while (upper >= lower) {
			at = (lower + upper) / 2;
			if (trove[x].get(at) == y) {
//				if (y == 101 && x ==1) {
//					System.out.println("found it");
//				}
//				if (!check.get(y, x)) {
//					Scanner s = new Scanner(System.in);
//					s.nextLine();
//				}
				return true;
			} else if (trove[x].get(at) > y) {
				upper = at - 1;
			} else if (trove[x].get(at) < y) {
				lower = at + 1;
			} else {
				System.out.println("something bad");
			}
		}
//		if (y == 101 && x ==1) {
//			System.out.println("did not find it");
//		}
//		if (check.get(y, x)) {
//			Scanner s = new Scanner(System.in);
//			s.nextLine();
//		}
		return false;
	}

	public void set(int y, int x, boolean in) {
		if (y < x) {
			int c = x;
			x = y;
			y = c;
		}

//		check.set(y, x, in);
//		if (y == 101 && x ==1) {
//			System.out.println("inserting " + x + " " + y + " " + in);
//			for (int i = 0; i < trove[x].size(); i++) {
//				System.out.print(trove[x].get(i) + " ");
//			}
//			System.out.println();
//		}
		if (x >= trove.length) {
			// grow the trove
			TIntList[] temp = new TIntArrayList[trove.length + GROWBY];
			for (int i = 0; i < trove.length; i++) {
				temp[i] = trove[i];
			}
			for (int i = trove.length; i < temp.length; i++) {
				temp[i] = new TIntArrayList();
			}
			trove = temp;
		}
		boolean go = true;
		boolean here = false;
		int upper = trove[x].size() - 1;
		int lower = 0;
		int at = (lower + upper) / 2;
		;
		// int counter=0;
		while (go && upper >= lower) {
			at = (lower + upper) / 2;
			if (trove[x].get(at) == y) {
				go = false;
				here = true;
			} else if (trove[x].get(at) > y) {
				upper = at - 1;
			} else if (trove[x].get(at) < y) {
				lower = at + 1;
			} else {
				System.out.println("something bad");
			}
		}
		if (lower > trove[x].size() - 1) {
			at = trove[x].size();
		}
		if (upper < 0) {
			at = 0;
		}
		if (here && !in) {
			trove[x].remove(y);
//			if (y == 101 && x ==1) {
//			System.out.println("i removed");
//			}
			}
		if (!here && in) {
			if (at > trove[x].size() - 1) {
				trove[x].add(y);
//				if (y == 101 && x ==1) {
//				System.out.println("i added");
//				}
			} else {
				trove[x].insert(at, y);
//				if (y == 101 && x ==1) {
//				System.out.println("i insereted");
//			}
				}
		}

//		if (get(x, y) != in) {
//			System.out.println("I suck");
//		}
//
//		if (check.get(y, x) != in) {
//			Scanner s = new Scanner(System.in);
//			s.nextLine();
//		}
	}

	public Straight getStraight(int x) {
		return straights.get(x);
	}

	public int value(int x) {
		return straights.get(x).value();
	}

	public int size() {
		return straights.size();
	}

	public void insert(Straight s) {
		straights.add(s);
	}
}
