package org.exoplatform.extension.exchange.model;

import java.io.Serializable;

public class FolderBean implements Serializable {
  private static final long serialVersionUID   = 4517749353533921356L;

  String                    id;

  String                    name;

  boolean                   synchronizedFolder = false;

  public FolderBean(String id, String name, boolean synchronizedFolder) {
    this.id = id;
    this.name = name;
    this.synchronizedFolder = synchronizedFolder;
  }

  public String getId() {
    return id;
  }

  public void setId(String id) {
    this.id = id;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public boolean isSynchronizedFolder() {
    return synchronizedFolder;
  }

  public void setSynchronizedFolder(boolean synchronizedFolder) {
    this.synchronizedFolder = synchronizedFolder;
  }
}
