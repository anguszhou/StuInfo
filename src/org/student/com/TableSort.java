package org.student.com;

import java.util.Comparator;

public class TableSort implements Comparator<Integer> {

	@Override
	public int compare(Integer o1, Integer o2) {
		// TODO Auto-generated method stub
		int i = o1 - o2;
		return -i;
	}

	

}
