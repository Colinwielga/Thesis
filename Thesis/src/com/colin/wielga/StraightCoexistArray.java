package com.colin.wielga;

import java.util.ArrayList;

public class StraightCoexistArray {
	int[][] ray;
	int size;
	private static final int GROWBY = 1000000;
	private static final int STARTMAX = 10000000;
	ArrayList<Straight> straights = new ArrayList<Straight>();
	public int doneUpTo;
	int currentmax= STARTMAX;
	
	public StraightCoexistArray() {
		ray = new int[STARTMAX][2];
		for (int i=0;i<STARTMAX;i++){
			ray[i][0] = -1;
			ray[i][1] = -1;
		}
		size = 0;
	}

	public boolean get(int y, int x) {
		if (y < x){
			int c = x;
			x = y;
			y=c;
		}
		int upper = size;
		int lower = 0;
		int at;
		//int counter=0;
		while (upper != lower){
//			counter++;
//			if (counter>100){
//				System.out.println("stuck in get "+upper+" "+lower);
//			}
			at = (lower + upper)/2;
			if (ray[at][0]==x && ray[at][1]==y){
				return true;
			}else if (ray[at][0] == -1||upper<lower){
				return false;
			}
			else if (ray[at][0]>x || (ray[at][0]==x &&ray[at][1]>y)){
				upper = at-1;
			}
			else if (ray[at][0]<x || (ray[at][0]==x &&ray[at][1]<y)){
				lower = at +1;
			}else{
				System.out.println("something bad");
			}
		}
		if (ray[upper][0]==x && ray[upper][1]==y){
			return true;
		}
		return false;
	}

	public Straight getStraight(int x) {
		return straights.get(x);
	}

	public void set(int y, int x, boolean in) {
		if (y < x){
			int c = x;
			x = y;
			y=c;
		}
		if (size+1 >= currentmax){
			int[][] temp = new int[currentmax + GROWBY][2]; 
			for (int i=0;i<currentmax;i++){
					temp[i][0] = ray[i][0];
					temp[i][1] = ray[i][1];
			}
			for (int i=currentmax;i<currentmax+ GROWBY;i++){
				temp[i][0] = -1;
				temp[i][1] = -1;
			}
			currentmax = currentmax + GROWBY;
			ray = temp;
			System.out.println("size is " + currentmax);
		}
		int upper = size;
		int lower = 0;
		int at;
		int counter =0;
		while (upper != lower){
			counter++;
			if (counter>100){
				System.out.println("stuck in set "+upper+" "+lower);
			}
			at = (lower + upper)/2;
			if (ray[at][0]==x && ray[at][1]==y){
				upper = at;
				lower =at;
			}else if (ray[at][0] == -1|| upper<lower){
				upper = at;
				lower = at;
			}
			else if (ray[at][0]>x || (ray[at][0]==x &&ray[at][1]>y)){
				upper = at-1;
			}
			else if (ray[at][0]<x || (ray[at][0]==x &&ray[at][1]<y)){
				lower = at +1;
			}else{
				System.out.println("something bad");
			}
		}
		if (ray[upper][0]==x && ray[upper][1]==y){
			if (!in){
				for (int i= upper;i<size-1;i++){
					ray[i] = ray[i+1];
				}
				ray[size][0]=-1;
				ray[size][1]=-1;
				size = size-1;
			}
		}else if (in){
			for (int i= upper;i<size+1;i++){
				ray[i+1] = ray[i];
			}
			ray[upper][0]=x;
			ray[upper][1]=y;
			size = size+1;
		}
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


	public void reset() {
		for (int i=0;i<size;i++){
				ray[i][0] = -1;
				ray[i][1] = -1;
		}
		size = 0;
		straights = new ArrayList<Straight>();
		doneUpTo =0;
	}
}
