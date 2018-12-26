package org.exoplatform.extension.exchange.listener;

import javax.jcr.Node;

import org.apache.commons.chain.Context;

import com.ibm.icu.util.Calendar;

import org.exoplatform.calendar.service.Utils;
import org.exoplatform.extension.exchange.service.IntegrationService;
import org.exoplatform.services.command.action.Action;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.ConversationState;
import org.exoplatform.services.security.IdentityConstants;

/**
 * @author Boubaker Khanfir
 */
@SuppressWarnings("all")
public class CalendarDeleteAction implements Action {

  private static final Log LOG = ExoLogger.getLogger(CalendarDeleteAction.class);

  public boolean execute(Context context) throws Exception {
    Node node = (Node) context.get("currentItem");
    if (node.isNodeType("exo:calendarEvent")) {
      String eventId = node.getName();
      try {
        String userId = null;
        ConversationState state = ConversationState.getCurrent();
        if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
          userId = node.getNode("../../../../..").getName();
        } else {
          userId = state.getIdentity().getUserId();
        }

        IntegrationService integrationService = IntegrationService.getInstance(userId);
        if (integrationService == null) {
          if (LOG.isTraceEnabled()) {
            LOG.info("User '" + (state == null || state.getIdentity() == null ? "Anonymous" : state.getIdentity().getUserId())
                + "' has no Exchange service, event will not be deleted from Exchange: eventId=" + eventId);
          }
          return false;
        } else {
          boolean started = false;
          try {
            String calendarId = node.getProperty(Utils.EXO_CALENDAR_ID).getString();
            if (integrationService.isCalendarSynchronizedWithExchange(calendarId)) {
              // Test if synchronization task is started, if yes, don't take
              // care about modifications to not corrupt data by cocurrent
              // modifications.
              if (!integrationService.isSynchronizationStarted()) {
                integrationService.setSynchronizationStarted();
                started = true;
                integrationService.deleteExchangeCalendarEvent(eventId, calendarId);
                integrationService.setUserExoLastCheckDate(Calendar.getInstance().getTime().getTime());
                integrationService.setSynchronizationStopped();
              }
            }
          } catch (Exception e) {
            LOG.error("Error while deleting Exchange event: " + eventId, e);
          } finally {
            // Set synchronization as finished if it was started here.
            if (started) {
              integrationService.setSynchronizationStopped();
            }
          }
        }
      } catch (Exception e) {
        LOG.error("Error while deleting Exchange item corresponding event to eXo Event with Id: " + eventId, e);
      }
    }
    return false;
  }
}
