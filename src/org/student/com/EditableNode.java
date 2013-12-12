package org.student.com;

import javax.swing.tree.DefaultMutableTreeNode;

public class EditableNode extends DefaultMutableTreeNode {
	
	private boolean isEdit;

	public EditableNode(Object userObject , boolean isEdit) {
		super(userObject);
		this.isEdit = isEdit;
	}

	public boolean isEdit() {
		return isEdit;
	}

	public void setEdit(boolean isEdit) {
		this.isEdit = isEdit;
	}
	
}
