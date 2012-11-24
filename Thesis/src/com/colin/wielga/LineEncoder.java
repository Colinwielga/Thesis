package com.colin.wielga;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Scanner;

public class LineEncoder {
	public static String DIMSINGLE;
	public static String PUBLICSINGLE;
	public static String DIMDOUBLE;
	public static String PUBLICDOUBLE;
	public static String DIMINTEGER;
	public static String PUBLICINTEGER;
	public static String DIMLONG;
	public static String PUBLICLONG;
	public static String DIMBOOLEAN;
	public static String PUBLICBOOLEAN;
	public static String DIMSTRING;
	public static String PUBLICSTRING;
	public static String DIMOTHER;
	public static String PUBLICOTHER;
	public static String STARTSUB;
	public static String ENDSUB;
	public static String STARTIF;
	public static String ENDIF;
	public static String STARTSELECT;
	public static String ENDSELECT;
	public static String CASE;
	public static String ELSE;
	public static String ELSEIF;
	public static String STARTFOR;
	public static String STARTDO;
	public static String STARTWHILE;
	public static String ENDFOR;
	public static String ENDDO;
	public static String ENDWHILE;
	public static String OPENFILEIN;
	public static String OPENFILEOUT;
	public static String FILEREAD;
	public static String FILEWRITE;
	public static String INPUTBOX;
	public static String MSGBOX;
	public static String PRINT;
	public static String ASSIGNMENT;
	public static String FORMLEVEL;
	public static String END;
	public static String STARTUNTIL;
	public static String STARTUNKNOWNDO;
	public static String ENDUNTIL;
	public static Scanner s = new Scanner(System.in);
	public static String OTHER;
	public static HashMap<String, String> hm = new HashMap<String, String>();
	private static String[] pullfrom ={"a","b","c","d","e","f","g","h","i","j","k","l","m","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
	private static int at=0;

	public static void load(String f) {
		if (Runner.debug) {
			System.out.println("loading encoding " + f);
		}
		String at = null;
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
					if (at == null) {
						if (strLine.equals("DIMSINGLE")
								|| strLine.equals("PUBLICSINGLE")
								|| strLine.equals("DIMDOUBLE")
								|| strLine.equals("PUBLICDOUBLE")
								|| strLine.equals("DIMINTEGER")
								|| strLine.equals("PUBLICINTEGER")
								|| strLine.equals("DIMLONG")
								|| strLine.equals("PUBLICLONG")
								|| strLine.equals("DIMBOOLEAN")
								|| strLine.equals("PUBLICBOOLEAN")
								|| strLine.equals("DIMSTRING")
								|| strLine.equals("PUBLICSTRING")
								|| strLine.equals("DIMOTHER")
								|| strLine.equals("PUBLICOTHER")
								|| strLine.equals("STARTSUB")
								|| strLine.equals("ENDSUB")
								|| strLine.equals("STARTIF")
								|| strLine.equals("ENDIF")
								|| strLine.equals("STARTSELECT")
								|| strLine.equals("ENDSELECT")
								|| strLine.equals("CASE")
								|| strLine.equals("ELSE")
								|| strLine.equals("STARTFOR")
								|| strLine.equals("STARTDO")
								|| strLine.equals("STARTWHILE")
								|| strLine.equals("STARTUNTIL")
								|| strLine.equals("ENDUNTIL")
								|| strLine.equals("ENDFOR")
								|| strLine.equals("ENDDO")
								|| strLine.equals("ENDWHILE")
								|| strLine.equals("OPENFILEIN")
								|| strLine.equals("OPENFILEOUT")
								|| strLine.equals("FILEREAD")
								|| strLine.equals("FILEWRITE")
								|| strLine.equals("INPUTBOX")
								|| strLine.equals("MSGBOX")
								|| strLine.equals("PRINT")
								|| strLine.equals("ASSIGNMENT")
								|| strLine.equals("FORMLEVEL")
								|| strLine.equals("END")
								|| strLine.equals("STARTUNITL")
								|| strLine.equals("STARTUNKNOWNDO")
								|| strLine.equals("ENDUNITL")) {
							at = new String(strLine);
						}
					} else {
						if (at.equals("DIMSINGLE")) {
							DIMSINGLE = strLine;
							at = null;
						} else if (at.equals("PUBLICSINGLE")) {
							PUBLICSINGLE = strLine;
							at = null;
						} else if (at.equals("DIMDOUBLE")) {
							DIMDOUBLE = strLine;
							at = null;
						} else if (at.equals("PUBLICDOUBLE")) {
							PUBLICDOUBLE = strLine;
							at = null;
						} else if (at.equals("DIMINTEGER")) {
							DIMINTEGER = strLine;
							at = null;
						} else if (at.equals("PUBLICINTEGER")) {
							PUBLICINTEGER = strLine;
							at = null;
						} else if (at.equals("DIMLONG")) {
							DIMLONG = strLine;
							at = null;
						} else if (at.equals("PUBLICLONG")) {
							PUBLICLONG = strLine;
							at = null;
						} else if (at.equals("DIMBOOLEAN")) {
							DIMBOOLEAN = strLine;
							at = null;
						} else if (at.equals("PUBLICBOOLEAN")) {
							PUBLICBOOLEAN = strLine;
							at = null;
						} else if (at.equals("DIMSTRING")) {
							DIMSTRING = strLine;
							at = null;
						} else if (at.equals("PUBLICSTRING")) {
							PUBLICSTRING = strLine;
							at = null;
						} else if (at.equals("DIMOTHER")) {
							DIMOTHER = strLine;
							at = null;
						} else if (at.equals("PUBLICOTHER")) {
							PUBLICOTHER = strLine;
							at = null;
						} else if (at.equals("STARTSUB")) {
							STARTSUB = strLine;
							at = null;
						} else if (at.equals("ENDSUB")) {
							ENDSUB = strLine;
							at = null;
						} else if (at.equals("STARTIF")) {
							STARTIF = strLine;
							at = null;
						} else if (at.equals("ENDIF")) {
							ENDIF = strLine;
							at = null;
						} else if (at.equals("STARTSELECT")) {
							STARTSELECT = strLine;
							at = null;
						} else if (at.equals("ENDSELECT")) {
							ENDSELECT = strLine;
							at = null;
						} else if (at.equals("CASE")) {
							CASE = strLine;
							at = null;
						} else if (at.equals("ELSE")) {
							ELSE = strLine;
							at = null;
						} else if (at.equals("ELSEIF")) {
							ELSEIF = strLine;
							at = null;
						} else if (at.equals("STARTFOR")) {
							STARTFOR = strLine;
							at = null;
						} else if (at.equals("STARTDO")) {
							STARTDO = strLine;
							at = null;
						} else if (at.equals("STARTWHILE")) {
							STARTWHILE = strLine;
							at = null;
						} else if (at.equals("ENDFOR")) {
							ENDFOR = strLine;
							at = null;
						} else if (at.equals("ENDDO")) {
							ENDDO = strLine;
							at = null;
						} else if (at.equals("ENDWHILE")) {
							ENDWHILE = strLine;
							at = null;
						} else if (at.equals("OPENFILEIN")) {
							OPENFILEIN = strLine;
							at = null;
						} else if (at.equals("OPENFILEOUT")) {
							OPENFILEOUT = strLine;
							at = null;
						} else if (at.equals("FILEREAD")) {
							FILEREAD = strLine;
							at = null;
						} else if (at.equals("FILEWRITE")) {
							FILEWRITE = strLine;
							at = null;
						} else if (at.equals("INPUTBOX")) {
							INPUTBOX = strLine;
							at = null;
						} else if (at.equals("MSGBOX")) {
							MSGBOX = strLine;
							at = null;
						} else if (at.equals("PRINT")) {
							PRINT = strLine;
							at = null;
						} else if (at.equals("ASSIGNMENT")) {
							ASSIGNMENT = strLine;
							at = null;
						} else if (at.equals("FORMLEVEL")) {
							FORMLEVEL = strLine;
							at = null;
						} else if (at.equals("END")) {
							END = strLine;
							at = null;
						} else if (at.equals("STARTUNTIL")) {
							STARTUNTIL = strLine;
							// System.out.print("STARTUNTIL is " + STARTUNTIL);
							at = null;
						} else if (at.equals("ENDUNTIL")) {
							ENDUNTIL = strLine;
							at = null;
						} else if (at.equals("STARTUNKNOWNDO")) {
							STARTUNKNOWNDO = strLine;
							at = null;
						} else if (at.equals("OTHER")) {
							OTHER = strLine;
							at = null;
						}

					}
				}

			}
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static String[] divide(String line) {
		ArrayList<String> result = new ArrayList<String>();
		line.trim();
		line.toLowerCase();
		boolean quotes = false;
		int lastcut = 0;
		for (int i = 0; i < line.length();) {
			if (line.charAt(i) == '"') {
				if (quotes) {
					// or i was going to cut but now that i think about it i
					// prolly don't even need to
				}
				quotes = !quotes;
				i++;
			} else if (!quotes) {
				if (line.charAt(i) == '\'') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					String[] realresult = new String[result.size()];
					result.toArray(realresult);
					return realresult;
				} else if (line.charAt(i) == '(') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add("(");
					i++;
					lastcut = i;
				} else if (line.charAt(i) == ')') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add(")");
					i++;
					lastcut = i;
				} else if (line.charAt(i) == ' ') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					i++;
					lastcut = i;
				} else if (line.charAt(i) == ',') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add(",");
					i++;
					lastcut = i;
				} else if (line.charAt(i) == '[') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add("[");
					i++;
					lastcut = i;
				} else if (line.charAt(i) == ']') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add("]");
					i++;
					lastcut = i;
				} else if (line.charAt(i) == ';') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add(";");
					i++;
					lastcut = i;
				} else // TODO what does ; do?
				if (line.charAt(i) == '=') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add("=");
					i++;
					lastcut = i;
				} else if (line.charAt(i) == '.') {
					if (!line.substring(lastcut, i).equals("")) {
						result.add(line.substring(lastcut, i));
					}
					result.add(".");
					i++;
					lastcut = i;
				} else {
					i++;
				}
			} else {
				i++;
			}

		}
		if (!line.substring(lastcut, line.length()).equals("")) {
			result.add(line.substring(lastcut, line.length()));
		}
		String[] realresult = new String[result.size()];
		result.toArray(realresult);
		return realresult;
	}

	public static String[] encode(String f) {
		//System.out.println("encodig file " + f);
		ArrayList<String> holder = new ArrayList<String>();
		File file = new File(f);
		FileInputStream fstream;
		try {
			fstream = new FileInputStream(file);
			// Get the object of DataInputStream
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine = null;
			while (strLine == null || strLine.trim() == "") {
				strLine = br.readLine();
			}
			// System.out.println("got here");
			strLine = strLine.trim().toLowerCase();
			// System.out.println(strLine);
			if (strLine.startsWith("version")) {
				// System.out.println("found version");
				boolean xml = true;
				int depth = 0;
				while (xml) {
					strLine = br.readLine();
					strLine = strLine.trim().toLowerCase();
					if (strLine.startsWith("begin")) {
						depth++;
						// System.out.println("found a begin");
					}
					if (strLine.startsWith("end")) {
						depth--;
						// System.out.println("found a end");
						if (depth == 0) {
							xml = false;
						}
					}

				}
			}
			while (br.ready()) {
				strLine = br.readLine();
				if (strLine != null && !strLine.equals("")) {
					// System.out.println("press enter");
					// s.nextLine();
					strLine = strLine.trim().toLowerCase();
					// TODO check to see if we need to do that " _" thing also
					// that
					// ; (:?) thing...

					String[] split = divide(strLine);
					if (split.length > 0) {
//						 System.out.println("");
//						 for (int i = 0; i < split.length; i++) {
//						 System.out.print(split[i] + ",");
//						 }
//						 System.out.println(split.length);
						// System.out.print("found: ");
						// DIM
						if (split[0].equals("dim")) {
							boolean toadd = true;
							for (int i = 0; i < split.length; i++) {
								// DIMSINGLE
								if (split[i].equals("single")) {
									holder.add(DIMSINGLE);
									// System.out.print(" " + DIMSINGLE);
									toadd = false;
								}
								// DIMDOUBLE
								if (split[i].equals("double")) {
									holder.add(DIMDOUBLE);
									// System.out.print(" " + DIMDOUBLE);
									toadd = false;
								}
								// DIMINTEGER
								if (split[i].equals("integer")) {
									holder.add(DIMINTEGER);
									// System.out.print(" " + DIMINTEGER);
									toadd = false;
								}
								// DIMLONG
								if (split[i].equals("long")) {
									holder.add(DIMLONG);
									// System.out.print(" " + DIMLONG);
									toadd = false;
								}
								// DIMBOOLEAN
								if (split[i].equals("boolean")) {
									holder.add(DIMBOOLEAN);
									// System.out.print(" " + DIMDOUBLE);
									toadd = false;
								}
								// DIMSTRING
								if (split[i].equals("string")) {
									holder.add(DIMSTRING);
									// System.out.print(" " + DIMSTRING);
									toadd = false;
								}
								// DIMOTHER
								if (split[i].equals(",")) {
									if (toadd) {
										// DIMOTHER
										holder.add(DIMOTHER);
										// System.out.print(" " + DIMOTHER);
									} else {
										toadd = false;
									}
								}
							}

						}
						if (split[0].equals("public")) {
							boolean toadd = true;
							for (int i = 0; i < split.length; i++) {
								// PUBLICSINGLE
								if (split[i].equals("single")) {
									holder.add(PUBLICSINGLE);
									// System.out.print(" " + PUBLICSINGLE);
									toadd = false;
								}
								// PUBLICDOUBLE
								if (split[i].equals("double")) {
									holder.add(PUBLICDOUBLE);
									// System.out.print(" " + PUBLICDOUBLE);
									toadd = false;
								}
								// PUBLICINTEGER
								if (split[i].equals("integer")) {
									holder.add(PUBLICINTEGER);
									// System.out.print(" " + PUBLICINTEGER);
									toadd = false;
								}
								// PUBLICLONG
								if (split[i].equals("long")) {
									holder.add(PUBLICLONG);
									// System.out.print(" " + PUBLICLONG);
									toadd = false;
								}
								// PUBLICBOOLEAN
								if (split[i].equals("boolean")) {
									holder.add(PUBLICBOOLEAN);
									// System.out.print(" " + PUBLICBOOLEAN);
									toadd = false;
								}
								// PUBLICSTRING
								if (split[i].equals("string")) {
									holder.add(PUBLICSTRING);
									// System.out.print(" " + PUBLICSTRING);
									toadd = false;
								}
								// PUBLICOTHER
								if (split[i].equals(",")) {
									if (toadd) {
										// PUBLICOTHER
										holder.add(PUBLICOTHER);
										// System.out.print(" " + PUBLICOTHER);
									} else {
										toadd = false;
									}
								}
							}

						}

						// STARTSUB
						if (split.length >= 2) {
							if (split[1].equals("sub")
									&& !split[0].equals("end")) {// and
																	// len
																	// >2?
								holder.add(STARTSUB);
								// System.out.print(" " + STARTSUB);
							}
							// ENDSUB
							if (split[1].equals("sub")
									&& split[0].equals("end")) {// and
																// len
																// =2?
								holder.add(ENDSUB);
								// System.out.print(" " + ENDSUB);
							}
							// ENDSELECT
							if (split[0].equals("end")
									&& split[1].equals("select")) {
								holder.add(ENDSELECT);
								// System.out.print(" " + ENDSELECT);
							}
							// ENDIF
							if (split[0].equals("end") && split[1].equals("if")) {
								holder.add(ENDIF);
								// System.out.print(" " + ENDIF);
							}
							// FILEWRITE
							if (split[0].equals("print") || isFile(split[1])) {
								// FILEWRITE
								holder.add(FILEWRITE);
								// System.out.print(" " + FILEWRITE);
							}
						}
						// STARTIF
						if (split[0].equals("if")) {
							holder.add(STARTIF);
							// System.out.print(" " + STARTIF);
						}

						// STARTSELECT
						if (split[0].equals("select")
								&& split[1].equals("case")) {
							holder.add(STARTSELECT);
							// System.out.print(" " + STARTSELECT);
						}
						// CASE
						if (split[0].equals("case")) {
							holder.add(CASE);
							// System.out.print(" " + CASE);
						}
						// ELSE
						if (split[0].equals("else")) {
							holder.add(ELSE);
							// System.out.print(" " + ELSE);
						}
						// ELSE
						if (split[0].equals("elseif")) {
							holder.add(ELSE);
							// System.out.print(" " + ELSEIF);
						}
						// STARTFOR
						if (split[0].equals("for")) {
							holder.add(STARTFOR);
							// System.out.print(" " + STARTFOR);
						}

						if (split[0].equals("do")) {
							if (split.length >= 2) {
								if (split[1].equals("while")) {
									// STARTWHILE
									holder.add(STARTWHILE);
									// System.out.print(" " + STARTWHILE);
								} else if (split[1].equals("until")) {
									// TODO add UNTIL to our list of shit
									// STARTUNTIL
									holder.add(STARTUNTIL);
									// System.out.print(" " + STARTUNTIL);
								} else {
									// UNKNOWDO
									holder.add(STARTUNKNOWNDO);
									// System.out.print(" " + STARTUNKNOWNDO);
									// holder.add(STARTDO);
								}
							} else {
								// UNKNOWDO
								holder.add(STARTUNKNOWNDO);
								// System.out.print(" " + STARTUNKNOWNDO);
								// holder.add(STARTDO);
							}
						}
						// ENDFOR
						if (split[0].equals("next")) {
							holder.add(ENDFOR);
							// System.out.print(" " + ENDFOR);
						}
						if (split[0].equals("loop")) {
							if (split.length == 1) {
								// go find the last unended loop and close it
								int loops = 0;
								int back = 1;
								while (!(loops == -1) && back <= holder.size()) {
									if (holder.get(holder.size() - back)
											.equals(STARTDO)
											// || holder.get(holder.size() -
											// back).equals(
											// STARTFOR)
											|| holder.get(holder.size() - back)
													.equals(STARTUNTIL)
											|| holder.get(holder.size() - back)
													.equals(STARTWHILE)
											|| holder.get(holder.size() - back)
													.equals(STARTUNKNOWNDO)) {
										loops--;
									}
									if (holder.get(holder.size() - back)
											.equals(ENDDO)
											|| holder.get(holder.size() - back)
													.equals(ENDUNTIL)
											// || holder.get(holder.size() -
											// back).equals(
											// ENDFOR)
											|| holder.get(holder.size() - back)
													.equals(ENDWHILE)) {
										loops++;
									}
									back++;
								}
								// TODO if back < holder.size shit is bad
								if (holder.get(holder.size() - back).equals(
										STARTDO)) {
									// ENDDO
									holder.add(ENDDO);
									// System.out.print(" " + ENDDO);
								}
								// else if (holder.get(holder.size() -
								// back).equals(
								// STARTFOR)) {
								// holder.add(ENDFOR);
								// }
								else if (holder.get(holder.size() - back)
										.equals(STARTUNTIL)) {
									holder.add(ENDUNTIL);
									// System.out.print(" " + ENDUNTIL);
								} else if (holder.get(holder.size() - back)
										.equals(STARTWHILE)) {
									holder.add(ENDWHILE);
									// System.out.print(" " + ENDWHILE);
								} else if (holder.get(holder.size() - back)
										.equals(STARTUNKNOWNDO)) {
									holder.set(holder.size() - back, STARTDO);
									// ENDDO
									holder.add(ENDDO);
									// System.out.print(" " + ENDDO);
								}
							} else {
								// TODO should a while with the condition before
								// be
								// a
								// different thing then a while with the
								// condition
								// after
								if (split[1].equals("until")) {
									boolean go = true;
									int back = 1;
									while (go) {
										if (holder.get(holder.size() - back)
												.equals(STARTUNKNOWNDO)) {
											go = false;
											holder.set(holder.size() - back,
													STARTUNTIL);
											holder.add(ENDUNTIL);
											// System.out.print(" " + ENDUNTIL);
										} else {
											back++;
										}
									}
								} else if (split[1].equals("while")) {
									boolean go = true;
									int back = 1;
									while (go) {
										if (holder.get(holder.size() - back)
												.equals(STARTUNKNOWNDO)) {
											go = false;
											holder.set(holder.size() - back,
													STARTWHILE);
											holder.add(ENDWHILE);
											// System.out.print(" " + ENDWHILE);
										} else {
											back++;
										}
									}
								}
							}
						}
						if (split[0].equals("open")) {
							boolean mode = false;
							for (int i = 1; i < split.length; i++) {
								if (split[i].equals("input")) {
									// OPENFILEIN
									holder.add(OPENFILEIN);
									// System.out.print(" " + OPENFILEIN);
									mode = true;
								}
								if (split[i].equals("output")) {
									// OPENFILEOUT
									holder.add(OPENFILEOUT);
									// System.out.print(" " + OPENFILEOUT);
									mode = true;
								}
							}
							if (!mode) {
								// OPENFILEOTHER
								// TODO i think this needs to be a thing
							}
						}
						// FILEREAD
						// TODO i am not very sure on this
						if (split[0].equals("read") || split[0].equals("input")
								|| split[0].equals("get")) { // i think...
							// FILEREAD
							holder.add(FILEREAD);
							// System.out.print(" " + FILEREAD);
						}
						for (int i = 0; i < split.length; i++) {
							if (split[i].equals("inputbox")) {
								// INPUTBOX
								holder.add(INPUTBOX);
								// System.out.print(" " + INPUTBOX);
							}
							if (split[i].equals("msgbox")) {
								// MSGBOX
								holder.add(MSGBOX);
								// System.out.print(" " + MSGBOX);
							}
							if (split[i].equals("print")) {
								// PRINT
								holder.add(PRINT);
								// System.out.print(" " + PRINT);
							}
							if (split[i].equals("=")
									&& (!split[0].equals("if"))
									&& (!split[0].equals("elseif"))
									&& (!split[0].equals("case"))) {
								// ASSIGNMENT
								holder.add(ASSIGNMENT);
								// System.out.print(" " + ASSIGNMENT);
							}
						}
						// FORMLEVEL
						// END
						if (split[0].equals("end") && split.length == 1) {
							holder.add(END);
							//System.out.print(" " + END);
						}
					}
				}
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
		String[] result = new String[holder.size()];
		holder.toArray(result);
		return result;
	}

	private static boolean isFile(String string) {
		String end = string.substring(1);
		if (string.startsWith("#")){
			try{
				Float.parseFloat(end);
				return true;
			}finally{
				return false;
			}
		}
		return false;
	}

	public static String tostr(String[] in) {
		String result = "";
		for (int i = 0; i < in.length; i++) {
			if (hm.containsKey(in[i])) {
				result = result +hm.get(in[i]);
			}else{
				hm.put(in[i], pullfrom[at]);
				at++;
				result = result +hm.get(in[i]);
			}
		}
		return result;
	}

}
