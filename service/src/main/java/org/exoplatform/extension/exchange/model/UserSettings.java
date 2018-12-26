package org.exoplatform.extension.exchange.model;

import java.io.Serializable;

public class UserSettings implements Serializable {
  private static final long serialVersionUID = -3248503274980906631L;

  private String            serverName;

  private String            domainName;

  private String            username;

  private String            password;

  public String getServerName() {
    return serverName;
  }

  public void setServerName(String serverName) {
    this.serverName = serverName;
  }

  public String getDomainName() {
    return domainName;
  }

  public void setDomainName(String domainName) {
    this.domainName = domainName;
  }

  public String getUsername() {
    return username;
  }

  public void setUsername(String username) {
    this.username = username;
  }

  public String getPassword() {
    return password;
  }

  public void setPassword(String password) {
    this.password = password;
  }

}
