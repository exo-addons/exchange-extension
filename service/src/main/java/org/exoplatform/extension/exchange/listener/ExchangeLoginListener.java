package org.exoplatform.extension.exchange.listener;

import org.apache.commons.lang.StringUtils;

import org.exoplatform.extension.exchange.service.SynchronizationService;
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

  private SynchronizationService synchronizationService;

  public ExchangeLoginListener(SynchronizationService synchronizationService) {
    this.synchronizationService = synchronizationService;
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
          synchronizationService.userLoggedOut(username);
        } catch (Exception e) {
          LOG.error("Error while user logout from MS Exchange Server", e);
        }
      }
    } else {
      // Login
      String username = event.getData().getIdentity().getUserId();
      if (StringUtils.isNotBlank(username) && !username.equals(IdentityConstants.ANONIM)) {
        try {
          synchronizationService.userLoggedIn(username, null);
        } catch (Exception e) {
          LOG.error("Error while authenticating user to MS Exchange Server", e);
        }
      }
    }
  }

}
