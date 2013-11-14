package org.exoplatform.extension.exchange.listener;

import javax.security.auth.callback.Callback;
import javax.security.auth.callback.NameCallback;
import javax.security.auth.callback.PasswordCallback;
import javax.security.auth.login.LoginException;

import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.jaas.AbstractLoginModule;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class ExchangeIntegrationLoginModule extends AbstractLoginModule {

  private static final Log LOG = ExoLogger.getLogger(ExchangeIntegrationLoginModule.class);

  private IntegrationListener exchangeListenerService;
  private String username = null;

  public ExchangeIntegrationLoginModule() {
    super();
  }

  @Override
  public boolean commit() throws LoginException {
    Callback[] callbacks = new Callback[2];
    callbacks[0] = new NameCallback("Username");
    callbacks[1] = new PasswordCallback("Password", false);
    try {
      callbackHandler.handle(callbacks);
      username = ((NameCallback) callbacks[0]).getName();
      String password = new String(((PasswordCallback) callbacks[1]).getPassword());
      if (username == null || username.isEmpty()) {
        // Let other login modules handle this
        return true;
      }
      if (password == null || password.isEmpty()) {
        // Let other login modules handle this
        return true;
      }
      getExchangeListenerService().userLoggedIn(username, password);
    } catch (Exception e) {
      e.printStackTrace();
      getLogger().error(e);
    }
    // Let other login modules run
    return true;
  }

  @Override
  public boolean login() throws LoginException {
    return true;
  }

  @Override
  public boolean abort() throws LoginException {
    return true;
  }

  @Override
  public boolean logout() throws LoginException {
    if (username != null) {
      exchangeListenerService.userLoggedOut(username);
    }
    return false;
  }

  @Override
  protected Log getLogger() {
    return LOG;
  }

  public IntegrationListener getExchangeListenerService() {
    if (exchangeListenerService == null) {
      try {
        this.exchangeListenerService = (IntegrationListener) getContainer().getComponentInstanceOfType(IntegrationListener.class);
      } catch (Exception e) {
        LOG.error(e);
      }
    }
    return exchangeListenerService;
  }
}
