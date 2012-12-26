package com.colin.wielga;

import java.util.ArrayList;

//TODO if we try to get or set bigger than the matrix we need to expand the matrix .. and that sucks and i need to write code to do it since i have a starting size of n^2 X n^2 and it needs to be 2^n X 2^n

import java.util.HashMap;

public class StraightCoexistMat {
	private static final int GROWBY = 1000;
	private static final int STARTMAX = 10000;
	boolean[][] mat;
	// HashMap<Integer, HashMap<Integer, Boolean>> mat = new HashMap<Integer,
	// HashMap<Integer, Boolean>>();
	ArrayList<Straight> straights = new ArrayList<Straight>();
	public int doneUpTo;
	int currentmax= STARTMAX;;

	public StraightCoexistMat() {
		mat = new boolean[STARTMAX][STARTMAX];
	}

	public boolean get(int y, int x) {
		// if (mat.containsKey(y)){
		// if (mat.get(y).containsKey(x)){
		// return mat.get(y).get(x);
		// }
		// }
		// return false;
		return mat[y][x];
	}

	public Straight getStraight(int x) {
		return straights.get(x);
	}

	public void set(int y, int x, Boolean in) {
		// if (mat.containsKey(y)){
		// mat.put(key, value)
		// }
		mat[y][x] = in;
	}

	public int value(int x) {
		return straights.get(x).value();
	}

	public int size() {
		return straights.size();
	}

	public void insert(Straight s) {
		straights.add(s);
		if (straights.size() >= currentmax){
			boolean[][] temp = new boolean[currentmax + GROWBY][currentmax + GROWBY]; 
			for (int i=0;i<currentmax;i++){
				for (int j =0; j<currentmax;j++){
					temp[i][j] = mat[i][j];
				}
			}
			currentmax = currentmax + GROWBY;
			mat = temp;
			System.out.println("size is " + currentmax);
		}
		
	}

	public void print() {
		for (int i = 0; i < straights.size(); i++) {
			for (int j = 0; j < straights.size(); j++) {
				System.out.print(mat[i][j] + " ");
			}
			System.out.println();
		}
	}

	public void reset() {
		for (int i=0;i<straights.size();i++){
			for (int j =0; j<straights.size();j++){
				mat[i][j] = false;
			}
		}
		straights = new ArrayList<Straight>();
		doneUpTo =0;
	}
}

// hmmm i think this needs to not be a hashmap.. i dont like hash maps this
// shoulod just be boolean[][] and them have an idea of size