package com.colin.wielga;

import java.util.ArrayList;

public class StraightCoexistArrayThing {
	ArrayList ray;
	int size;
	private static final int GROWBY = 1000000;
	private static final int STARTMAX = GROWBY;
	ArrayList<Straight> straights = new ArrayList<Straight>();
	public int doneUpTo;
	//int currentmax= STARTMAX;
	
	public StraightCoexistArrayThing() {
		ray = new ArrayList();
		int[][] temp =new int[STARTMAX][2];
		for (int i=0;i<STARTMAX;i++){
			temp[i][0] = -1;
			temp[i][1] = -1;
		}
		ray.add(temp);
		size = 0;
	}
	
	public void rayset(int x,int y,int value){
		((int[][]) ray.get(x/GROWBY))[x%GROWBY][y] = value;
	}
	
	public int rayget(int x,int y){
		return ((int[][]) ray.get(x/GROWBY))[x%GROWBY][y];
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
			if (rayget(at,0)==x && rayget(at,1)==y){
				return true;
			}else if (rayget(at,0) == -1||upper<lower){
				return false;
			}
			else if (rayget(at,0)>x || (rayget(at,0)==x &&rayget(at,1)>y)){
				upper = at-1;
			}
			else if (rayget(at,0)<x || (rayget(at,0)==x &&rayget(at,1)<y)){
				lower = at +1;
			}else{
				System.out.println("something bad");
			}
		}
		if (rayget(upper,0)==x && rayget(upper,1)==y){
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
		if (size+1 >= ray.size()*GROWBY){
			int[][] temp = new int[GROWBY][2];
			for (int i=0;i<temp.length;i++){
				temp[i][0]=-1;
				temp[i][1] = -1;
			}
			ray.add(temp);
			//currentmax = currentmax + GROWBY;
			//System.out.println("size is " + currentmax);
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
			if (rayget(at,0)==x && rayget(at,1)==y){
				upper = at;
				lower =at;
			}else if (rayget(at,0) == -1|| upper<lower){
				upper = at;
				lower = at;
			}
			else if (rayget(at,0)>x || (rayget(at,0)==x &&rayget(at,1)>y)){
				upper = at-1;
			}
			else if (rayget(at,0)<x || (rayget(at,0)==x &&rayget(at,1)<y)){
				lower = at +1;
			}else{
				System.out.println("something bad");
			}
		}
		if (rayget(upper,0)==x && rayget(upper,1)==y){
			if (!in){
				for (int i= upper;i<size-1;i++){
					rayset(i,0,rayget(i+1,0));
					rayset(i,1,rayget(i+1,1));
				}
				rayset(size,0,-1);
				rayset(size,1,-1);
				size = size-1;
			}
		}else if (in){
			for (int i= upper;i<size+1;i++){//once was size +1
				rayset(i+1,0,rayget(i,0));
				rayset(i+1,1,rayget(i,1));
			}
			rayset(upper,0,x);
			rayset(upper,1,y);
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
		ray = new ArrayList();
		int[][] temp =new int[STARTMAX][2];
		for (int i=0;i<STARTMAX;i++){
			temp[i][0] = -1;
			temp[i][1] = -1;
		}
		ray.add(temp);
		size = 0;
		straights = new ArrayList<Straight>();
		doneUpTo =0;
	}


}
