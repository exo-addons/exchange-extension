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

/**
 * @author Boubaker Khanfir
 */
@SuppressWarnings("all")
public class CalendarCreateUpdateAction implements Action {

  public final static ThreadLocal<Long>     MODIFIED_DATE         = new ThreadLocal<Long>();

  private final static ThreadLocal<Boolean> IGNORE_UPDATE         = new ThreadLocal<Boolean>();

  private static final String               EXO_DATETIME_PROPERTY = "exo:datetime";

  private static final Log                  LOG                   = ExoLogger.getLogger(CalendarCreateUpdateAction.class);

  public boolean execute(Context context) throws Exception {
    if (IGNORE_UPDATE.get() != null && IGNORE_UPDATE.get()) {
      return false;
    }
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
    if (!isNodeValid(node)) {
      return false;
    }

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
        LOG.warn("No authenticated user was found while trying to create/update eXo Calendar event with id: '" + eventId
            + "' for user: " + userId);
        return false;
      } else if (MODIFIED_DATE.get() == null || MODIFIED_DATE.get() == 0) {
        try {
          modifyUpdateDate(node, System.currentTimeMillis());
          integrationService.updateOrCreateExchangeCalendarEvent(node);
        } catch (Exception e) {
          LOG.warn("Error while create/update an Exchange item for eXo event: " + eventId, e);
        }
      } else {
        modifyUpdateDate(node, MODIFIED_DATE.get());
      }
    } catch (Exception e) {
      LOG.error("Error while updating Exchange with the eXo Event with Id: " + eventId, e);
    }
    return false;
  }

  private void modifyUpdateDate(Node node, long lastModifiedDate) throws Exception {
    IGNORE_UPDATE.set(true);
    try {
      GregorianCalendar modifiedDate = new GregorianCalendar();
      if (lastModifiedDate > 0) {
        modifiedDate.setTimeInMillis(lastModifiedDate);
      }

      if (!node.isNodeType(EXO_DATETIME_PROPERTY)) {
        if (node.canAddMixin(EXO_DATETIME_PROPERTY)) {
          node.addMixin(EXO_DATETIME_PROPERTY);
        }
        node.setProperty("exo:dateCreated", modifiedDate);
      }
      node.setProperty("exo:dateModified", modifiedDate);
    } finally {
      IGNORE_UPDATE.set(false);
    }
  }

  private boolean isNotLastPropertyToSet(Node node, Property property) throws Exception {
    return (property.getName().equals(Utils.EXO_PARTICIPANT_STATUS)
        && (!node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT) || (node.hasProperty(Utils.EXO_REPEAT)
            && node.getProperty(Utils.EXO_REPEAT).getString().equals(CalendarEvent.RP_NOREPEAT))));
  }

  private boolean isNodeValid(Node node) throws Exception {
    return node != null && node.isNodeType("exo:calendarEvent")
        && ((node.hasProperty(Utils.EXO_PARTICIPANT_STATUS) && !node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT))
            || (node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT) && node.hasProperty(Utils.EXO_REPEAT_INTERVAL)));
  }

}
