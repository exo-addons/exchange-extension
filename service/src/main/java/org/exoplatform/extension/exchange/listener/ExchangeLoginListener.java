package org.exoplatform.extension.exchange.listener;

import org.apache.commons.lang.StringUtils;

import org.exoplatform.services.listener.*;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.*;

/**
 * @author Boubaker Khanfir
 */
@Asynchronous
public class ExchangeLoginListener extends Listener<ConversationRegistry, ConversationState> {

  private static final Log    LOG = ExoLogger.getLogger(ExchangeLoginListener.class);

  private IntegrationListener exchangeListenerService;

  public ExchangeLoginListener(IntegrationListener integrationListener) {
    this.exchangeListenerService = integrationListener;
  }

  @Override
  public void onEvent(Event<ConversationRegistry, ConversationState> event) throws Exception {
    String eventName = event.getEventName();
    if (eventName.endsWith(".unregister")) {
      // Logout
      String username = event.getData() == null
          || event.getData().getIdentity() == null ? null : event.getData().getIdentity().getUserId();
      if (StringUtils.isNotBlank(username) && !username.equals(IdentityConstants.ANONIM)) {
        try {
          exchangeListenerService.userLoggedOut(username);
        } catch (Exception e) {
          LOG.error("Error while user logout from MS Exchange Server", e);
        }
      }
    } else {
      // Login
      String username = event.getData().getIdentity().getUserId();
      if (StringUtils.isNotBlank(username) && !username.equals(IdentityConstants.ANONIM)) {
        try {
          exchangeListenerService.userLoggedIn(username, null);
        } catch (Exception e) {
          LOG.error("Error while authenticating user to MS Exchange Server", e);
        }
      }
    }
  }

}
