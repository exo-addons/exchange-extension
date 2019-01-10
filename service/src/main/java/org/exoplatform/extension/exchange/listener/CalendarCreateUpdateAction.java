package org.exoplatform.extension.exchange.listener;

import java.util.*;

import javax.jcr.Node;
import javax.jcr.Property;

import org.apache.commons.chain.Context;
import org.apache.commons.lang.StringUtils;

import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.Utils;
import org.exoplatform.extension.exchange.task.UserIntegrationFacade;
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

  private final static List<String>         PROPERTIES_WATCHED    = Arrays.asList(new String[] { Utils.EXO_PARTICIPANT_STATUS,
      Utils.EXO_REPEAT_INTERVAL, Utils.EXO_REPEAT_FINISH_DATE, Utils.EXO_REPEAT_BYMONTHDAY, Utils.EXO_REPEAT_BYDAY });

  private static final String               EXO_DATETIME_PROPERTY = "exo:datetime";

  private static final Log                  LOG                   = ExoLogger.getLogger(CalendarCreateUpdateAction.class);

  public boolean execute(Context context) throws Exception {
    if (IGNORE_UPDATE.get() != null && IGNORE_UPDATE.get()) {
      return false;
    }
    Object object = context.get("currentItem");
    Node node = null;
    Property property = null;
    if (object instanceof Node) {
      node = (Node) object;
    } else if (object instanceof Property) {
      property = (Property) object;
      node = property.getParent();
    }
    if (!isNodeValid(node)) {
      return false;
    }

    // This is to avoid making multiple update requests to exchange
    if (property != null && !isLastPropertyToSet(node, property)) {
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

      UserIntegrationFacade integrationService = UserIntegrationFacade.getInstance(userId);
      if (integrationService == null) {
        LOG.trace("No authenticated user was found while trying to create/update eXo Calendar event with id: '{}' for user: {}",
                  eventId,
                  userId);
        return false;
      } else {
        Long lastModifiedDate = MODIFIED_DATE.get();
        if (lastModifiedDate == null || lastModifiedDate == 0) {
          String calendarId =
                            node.hasProperty(Utils.EXO_CALENDAR_ID) ? node.getProperty(Utils.EXO_CALENDAR_ID).getString() : null;
          if (integrationService.isCalendarSynchronizedWithExchange(calendarId)) {
            boolean started = integrationService.setSynchronizationStarted();
            if (started) {
              try {
                modifyUpdateDate(node, System.currentTimeMillis());
                integrationService.updateOrCreateExchangeCalendarEvent(node);
              } catch (Exception e) {
                LOG.warn("Error while create/update an Exchange item for eXo event: " + eventId, e);
              } finally {
                integrationService.setSynchronizationStopped();
              }
            }
          }
        } else {
          modifyUpdateDate(node, lastModifiedDate);
        }
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

  private boolean isLastPropertyToSet(Node node, Property property) throws Exception {
    String propertyName = property.getName();
    if (!PROPERTIES_WATCHED.contains(propertyName)) {
      return false;
    }

    if (!node.isNodeType(Utils.EXO_REPEAT_CALENDAR_EVENT)) {
      // EXO_PARTICIPANT_STATUS is last property set
      return true;
    } else {
      // Check if EXO_REPEAT_FINISH_DATE is the last property to set
      Calendar dateRepeatTo = node.hasProperty(Utils.EXO_REPEAT_UNTIL) && node.getProperty(Utils.EXO_REPEAT_UNTIL) != null
          && node.getProperty(Utils.EXO_REPEAT_UNTIL).getValue() != null ? node.getProperty(Utils.EXO_REPEAT_UNTIL).getDate()
                                                                         : null;
      if (dateRepeatTo != null) {
        return propertyName.equals(Utils.EXO_REPEAT_FINISH_DATE);
      }

      // Check if EXO_REPEAT_FINISH_DATE is the last property to set
      Long repeatCount = node.hasProperty(Utils.EXO_REPEAT_COUNT) && node.getProperty(Utils.EXO_REPEAT_COUNT) != null
          && node.getProperty(Utils.EXO_REPEAT_COUNT).getValue() != null ? node.getProperty(Utils.EXO_REPEAT_COUNT).getLong()
                                                                         : null;
      if (repeatCount != null && repeatCount > 0) {
        return propertyName.equals(Utils.EXO_REPEAT_FINISH_DATE);
      }

      // Check if EXO_REPEAT_BYDAY or EXO_REPEAT_BYMONTHDAY is the last property
      // to set
      String repeatType = node.hasProperty(Utils.EXO_REPEAT) && node.getProperty(Utils.EXO_REPEAT) != null
          && node.getProperty(Utils.EXO_REPEAT).getValue() != null ? node.getProperty(Utils.EXO_REPEAT).getString() : null;
      if (StringUtils.equals(repeatType, CalendarEvent.RP_WEEKLY)) {
        return propertyName.equals(Utils.EXO_REPEAT_BYDAY);
      } else if (StringUtils.equals(repeatType, CalendarEvent.RP_MONTHLY)) {
        return propertyName.equals(Utils.EXO_REPEAT_BYMONTHDAY);
      }

      // EXO_REPEAT_INTERVAL is the last property set
      return true;
    }
  }

  private boolean isNodeValid(Node node) throws Exception {
    return node != null && node.isNodeType(Utils.EXO_CALENDAR_EVENT);
  }

}
