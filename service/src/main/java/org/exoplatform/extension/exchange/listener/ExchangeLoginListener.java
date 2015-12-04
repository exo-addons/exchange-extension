package org.exoplatform.extension.exchange.listener;

import org.apache.commons.lang.StringUtils;
import org.exoplatform.container.PortalContainer;
import org.exoplatform.services.listener.Event;
import org.exoplatform.services.listener.Listener;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.ConversationRegistry;
import org.exoplatform.services.security.ConversationState;
import org.exoplatform.services.security.IdentityConstants;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class ExchangeLoginListener extends Listener<ConversationRegistry, ConversationState> {

  private static final Log LOG = ExoLogger.getLogger(ExchangeLoginListener.class);

  private IntegrationListener exchangeListenerService;

  public IntegrationListener getExchangeListenerService() {
    if (exchangeListenerService == null) {
      try {
        this.exchangeListenerService = (IntegrationListener) PortalContainer.getInstance().getComponentInstanceOfType(IntegrationListener.class);
      } catch (Exception e) {
        LOG.error(e);
      }
    }
    return exchangeListenerService;
  }

  @Override
  public void onEvent(Event<ConversationRegistry, ConversationState> event) throws Exception {
    String username = event.getData().getIdentity().getUserId();
    if (StringUtils.isNotEmpty(username) && !username.equals(IdentityConstants.ANONIM)) {
      getExchangeListenerService().userLoggedIn(username, null);
    }
  }
}
