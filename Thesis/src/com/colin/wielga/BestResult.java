package com.colin.wielga;

public class BestResult {
	boolean foundone = false;
	Straight winner= null;
	 Straight[] dissallows;
	
	public BestResult(Straight winner, Straight[] dissallows){
		foundone = true;
		this.winner = winner;
		this.dissallows = dissallows;
	}

}
