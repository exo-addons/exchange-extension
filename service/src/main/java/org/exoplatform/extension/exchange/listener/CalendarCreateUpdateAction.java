package org.exoplatform.extension.exchange.listener;

import java.util.GregorianCalendar;

import javax.jcr.Node;
import javax.jcr.Property;

import org.apache.commons.chain.Context;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.Utils;
import org.exoplatform.extension.exchange.service.IntegrationService;
import org.exoplatform.services.command.action.Action;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.security.ConversationState;
import org.exoplatform.services.security.IdentityConstants;

import com.ibm.icu.util.Calendar;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
public class CalendarCreateUpdateAction implements Action {

  private final static Log LOG = ExoLogger.getLogger(CalendarCreateUpdateAction.class);

  public boolean execute(Context context) throws Exception {
    Object object = context.get("currentItem");
    Node node = null;
    if (object instanceof Node) {
      node = (Node) object;
    } else if (object instanceof Property) {
      Property property = (Property) object;
      node = property.getParent();
      // This is to avoid making multiple update requests to exchange
      if (!isNotLastPropertyToSet(node, property)) {
        return false;
      }
    }
    if (node != null && node.isNodeType("exo:calendarEvent") && isNodeValid(node)) {
      String eventId = node.getName();
      try {
        String userId = null;
        ConversationState state = ConversationState.getCurrent();
        if (state == null || state.getIdentity() == null || state.getIdentity().getUserId().equals(IdentityConstants.ANONIM)) {
          userId = node.getNode("../../../../..").getName();
        } else {
          userId = state.getIdentity().getUserId();
        }

        if (userId == null) {
          LOG.warn("No user was found while trying to create/update eXo Calendar event with id: " + eventId);
          return false;
        }

        IntegrationService integrationService = IntegrationService.getInstance(userId);
        if (integrationService == null) {
          LOG.warn("No authenticated user was found while trying to create/update eXo Calendar event with id: '" + eventId + "' for user: " + userId);
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
                if (integrationService.getUserExoLastCheckDate() != null) {
                  integrationService.updateOrCreateExchangeCalendarEvent(node);
                  modifyUpdateDate(node);
                  integrationService.setUserExoLastCheckDate(Calendar.getInstance().getTime().getTime());
                }
                integrationService.setSynchronizationStopped();
              }
            }
          } catch (Exception e) {
            // This can happen if the node was newly created, so not all
            // properties are in the node
            LOG.error("Error while create/update an Exchange item for eXo event: " + eventId, e);
            // Integration is out of sync, so disable auto synchronization
            // until the scheduled job runs and try to fix this
            integrationService.setUserExoLastCheckDate(0);
          } finally {
            // Set synchronization as finished if it was started here.
            if (started) {
              integrationService.setSynchronizationStopped();
            }
          }
        }
      } catch (Exception e) {
        LOG.error("Error while updating Exchange with the eXo Event with Id: " + eventId, e);
      }
    }
    return false;
  }

  private void modifyUpdateDate(Node node) throws Exception {
    if (!node.isNodeType("exo:datetime")) {
      if (node.canAddMixin("exo:datetime")) {
        node.addMixin("exo:datetime");
      }
      node.setProperty("exo:dateCreated", new GregorianCalendar());
    }
    node.setProperty("exo:dateModified", new GregorianCalendar());
  }

  private boolean isNotLastPropertyToSet(Node node, Property property) throws Exception {
    return (property.getName().equals(Utils.EXO_PARTICIPANT_STATUS) && (!node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT) || (node.hasProperty(Utils.EXO_REPEAT) && node.getProperty(Utils.EXO_REPEAT)
        .getString().equals(CalendarEvent.RP_NOREPEAT))))
        || (node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT) && (property.getName().equals(Utils.EXO_REPEAT_BYMONTHDAY) || property.getName().equals(Utils.EXO_REPEAT_FINISH_DATE)));
  }

  private boolean isNodeValid(Node node) throws Exception {
    return (node.hasProperty(Utils.EXO_PARTICIPANT_STATUS) && !node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT))
        || (node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT) && node.hasProperty(Utils.EXO_REPEAT_INTERVAL));
  }

}
