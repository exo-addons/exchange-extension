package org.exoplatform.extension.exchange.listener;

import java.net.URI;
import java.util.*;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.extension.exchange.service.*;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.security.ConversationState;
import org.exoplatform.services.security.Identity;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.notification.EventType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.notification.*;
import microsoft.exchange.webservices.data.property.complex.FolderId;

/**
 * Thread used to synchronize Exchange Calendar with eXo Calendar
 */
@SuppressWarnings("deprecation")
public class ExchangeIntegrationTask extends Thread {
  private static final Log    LOG               = ExoLogger.getLogger(ExchangeIntegrationTask.class);

  private static final String LOG_SEPARATOR     = "---------------------";

  private ExchangeService     service;

  private PullSubscription    subscription      = null;

  private IntegrationService  integrationService;

  private List<FolderId>      calendarFolderIds = new ArrayList<>();

  private String              username;

  private ConversationState   state;

  private boolean             firstSynchronization;

  private boolean             synchronizeAllExchangeFolders;

  private boolean             deleteExoCalendarOnUnsync;

  public ExchangeIntegrationTask(OrganizationService organizationService,
                                 CalendarService calendarService,
                                 ExoStorageService exoStorageService,
                                 ExchangeStorageService exchangeStorageService,
                                 CorrespondenceService correspondenceService,
                                 Identity identity,
                                 String exchangeUsername,
                                 String exchangePassword,
                                 String exchangeDomain,
                                 String exchangeServerURL,
                                 boolean synchronizeAllExchangeFolders,
                                 boolean deleteExoCalendarOnUnsync)
      throws Exception {
    this.username = identity.getUserId();
    this.firstSynchronization = true;
    this.synchronizeAllExchangeFolders = synchronizeAllExchangeFolders;
    this.deleteExoCalendarOnUnsync = deleteExoCalendarOnUnsync;

    service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
    service.setTimeout(300000);

    ExchangeCredentials credentials = null;
    // Try with domain name if it's not empty
    if (exchangeDomain != null) {
      // Test authentication with
      // "exchangeUsername, exchangePassword, exchangeDomain"
      credentials = new WebCredentials(exchangeUsername, exchangePassword, exchangeDomain);
    } else {
      // Test authentication with "exchangeUsername, exchangePassword"
      credentials = new WebCredentials(exchangeUsername, exchangePassword);
    }
    service.setCredentials(credentials);
    service.setUrl(new URI(exchangeServerURL));

    try {
      service.getInboxRules();
    } catch (Exception e) {

      boolean authenticated = false;
      if (exchangeDomain != null) {
        // Test authentication with "exchangeUsername, exchangePassword" if
        // domainName not null
        credentials = new WebCredentials(exchangeUsername, exchangePassword);
        try {
          service.setCredentials(credentials);
          service.setUrl(new URI(exchangeServerURL));

          service.getInboxRules();
          authenticated = true;
        } catch (Exception exp) {
          if (!exchangeUsername.contains("@")) {
            // Test authentication with
            // "exchangeUsername@domainName, exchangePassword" if domainName
            // not null
            credentials = new WebCredentials(exchangeUsername + "@" + exchangeDomain, exchangePassword);
            try {
              service.setCredentials(credentials);
              service.setUrl(new URI(exchangeServerURL));

              service.getInboxRules();
              authenticated = true;
            } catch (Exception exp2) {
              authenticated = false;
            }

          }
        }
      }
      if (!authenticated && (exchangeDomain == null || exchangeDomain.isEmpty()) && exchangeUsername.contains("@")) {
        String[] parts = exchangeUsername.split("@");
        exchangeUsername = parts[0];
        exchangeDomain = parts[1];
        // Test authentication with
        // "exchangeUsername, exchangePassword" and domainName extracted
        // from
        // username
        credentials = new WebCredentials(exchangeUsername, exchangePassword, exchangeDomain);
        service.setCredentials(credentials);
        service.setUrl(new URI(exchangeServerURL));

        service.getInboxRules();
      } else {
        throw e;
      }
    }

    integrationService = new IntegrationService(organizationService,
                                                calendarService,
                                                exoStorageService,
                                                exchangeStorageService,
                                                correspondenceService,
                                                service,
                                                username);

    // Set current identity visible in this Thread
    state = new ConversationState(identity);
    ConversationState.setCurrent(state);

    // First call to the service, this may fail because of wrong
    // credentials
    if (synchronizeAllExchangeFolders) {
      calendarFolderIds = exchangeStorageService.getAllExchangeCalendars(service);
    } else {
      // Test connection
      Folder folder =
                    integrationService.getExchangeCalendar(FolderId.getFolderIdFromWellKnownFolderName(WellKnownFolderName.Calendar));
      if (folder != null) {
        integrationService.setSynchronizationStarted();
        calendarFolderIds = integrationService.getSynchronizedExchangeCalendars();
        integrationService.setSynchronizationStopped();
      } else {
        throw new IllegalStateException("Error while authenticating user '" + username
            + "' to exchange, please make sure you are connected to the correct URL with correct credentials.");
      }
    }
  }

  @Override
  public void run() {
    waitOtherTasks();
    try {
      integrationService.setSynchronizationStarted();

      ConversationState.setCurrent(state);

      long newLastTimeCheck = System.currentTimeMillis();

      // Verify Exchange folders state with Exo Calendars state
      List<String> updatedExoEventIDs = integrationService.synchronizeExchangeFolderState(calendarFolderIds,
                                                                                          synchronizeAllExchangeFolders,
                                                                                          deleteExoCalendarOnUnsync);
      if (calendarFolderIds.isEmpty()) {
        return;
      }
      if (updatedExoEventIDs == null) {
        updatedExoEventIDs = new ArrayList<>();
      }
      Date lastSyncDate = integrationService.getUserLastCheckDate();
      // This is used once, when user login
      if (firstSynchronization) {
        if (LOG.isTraceEnabled()) {
          LOG.trace(LOG_SEPARATOR);
          LOG.trace("run first synchronization for user: " + username);
        }

        // Begin catching events from Exchange
        newSubscription();

        // Verify modifications made on folders
        synchronizeByModificationDate(lastSyncDate, updatedExoEventIDs);
        this.firstSynchronization = false;
      } else {
        if (LOG.isTraceEnabled()) {
          LOG.trace(LOG_SEPARATOR);
          LOG.trace("run scheduled synchronization for user: " + username);
        }
        // This is used in a scheduled task when the user session still alive
        GetEventsResults events = getEvents();
        if (synchronizeAllExchangeFolders) {
          synchronizeExchangeFolders(events, updatedExoEventIDs);
        }
        synchronizeExchangeApointments(events, lastSyncDate, updatedExoEventIDs);
        synchronizeByModificationDate(lastSyncDate, updatedExoEventIDs);
      }

      // Update date of last check in a user profile attribute
      integrationService.setUserLastCheckDate(newLastTimeCheck);

      if (LOG.isTraceEnabled()) {
        LOG.trace("Synchronization completed.");
        LOG.trace(LOG_SEPARATOR);
      }
    } catch (Exception e) {
      LOG.error("Error while synchronizing calndar entries.", e);
    } finally {
      integrationService.setSynchronizationStopped();
    }
  }

  @Override
  public void interrupt() {
    if (subscription != null) {
      try {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Thread interruption: unsubscribe user service:" + username);
        }
        subscription.unsubscribe();
      } catch (Exception e) {
        LOG.error("Thread interruption: Error while unsubscribe to thread of user:" + username);
      }
    }

    if (service != null) {
      try {
        service.close();
      } catch (Exception e) {
        LOG.error("Thread interruption: Error while closing ExchangeService for user:" + username, e);
      }
    }

    try {
      integrationService.removeInstance();
    } catch (Throwable e) {
      LOG.error("Error while inerrupting thread", e);
    }
    super.interrupt();
  }

  private GetEventsResults getEvents() throws Exception {
    GetEventsResults events = null;
    try {
      events = subscription.getEvents();
    } catch (Exception e) {
      LOG.warn("Subscription seems timed out, retry. Original cause: " + e.getMessage() + "");
      newSubscription();
      events = subscription.getEvents();
    }
    return events;
  }

  private void waitOtherTasks() {
    int i = 0;
    while (integrationService.isSynchronizationStarted() && i < 5) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Exchange integration is in use, scheduled job will wait until synchronization is finished for user:'"
            + username + "'.");
      }
      try {
        Thread.sleep(1000);
      } catch (Exception e) {
        LOG.warn(e.getMessage());
      }
      i++;
    }
  }

  private void synchronizeByModificationDate(Date lastSyncDate, List<String> updatedExoEventIDs) throws Exception {
    // synchronize eXo Calendar with Exchange
    for (FolderId folderId : calendarFolderIds) {
      Calendar calendar = integrationService.getUserCalendarByExchangeFolderId(folderId);
      if (calendar == null || lastSyncDate == null) {
        integrationService.synchronizeFullCalendar(folderId);
      } else {
        integrationService.synchronizeModificationsOfCalendar(folderId, lastSyncDate, updatedExoEventIDs);
      }
    }
  }

  @SuppressWarnings("all")
  private long synchronizeExchangeApointments(GetEventsResults events,
                                              Date lastSyncDate,
                                              List<String> updatedExoEventIDs) throws Exception {
    // loop through Appointment events
    Iterable<ItemEvent> itemEvents = events.getItemEvents();
    long lastTimeCheck = System.currentTimeMillis();
    if (itemEvents.iterator().hasNext()) {
      List<String> itemIds = new ArrayList<>();
      for (ItemEvent itemEvent : itemEvents) {
        if (itemIds.contains(itemEvent.getItemId().getUniqueId())) {
          continue;
        }
        itemIds.add(itemEvent.getItemId().getUniqueId());
        List<CalendarEvent> updatedEvents = integrationService.createOrUpdateOrDelete(itemEvent, lastSyncDate);
        if (updatedEvents != null && !updatedEvents.isEmpty() && updatedExoEventIDs != null) {
          for (CalendarEvent calendarEvent : updatedEvents) {
            updatedExoEventIDs.add(calendarEvent.getId());
          }
        }
      }
    }
    return lastTimeCheck;
  }

  private void synchronizeExchangeFolders(GetEventsResults events, List<String> updatedExoEventIDs) throws Exception {
    // If Calendar Folders was modified
    if (events.getFolderEvents() != null && events.getFolderEvents().iterator().hasNext()) {
      Iterator<FolderEvent> iterator = events.getFolderEvents().iterator();
      while (iterator.hasNext()) {
        FolderEvent folderEvent = iterator.next();
        if (folderEvent.getEventType().equals(EventType.Created) || folderEvent.getEventType().equals(EventType.Modified)) {
          if (!integrationService.isCalendarPresentInExo(folderEvent.getFolderId())) {
            List<String> updatedEventIDs = integrationService.synchronizeFullCalendar(folderEvent.getFolderId());
            updatedExoEventIDs.addAll(updatedEventIDs);
            if (!updatedEventIDs.isEmpty() && !calendarFolderIds.contains(folderEvent.getFolderId())) {
              calendarFolderIds.add(folderEvent.getFolderId());
            }
          }
        } else if (folderEvent.getEventType().equals(EventType.Deleted)) {
          boolean deleted = integrationService.deleteExoCalendar(folderEvent.getFolderId());
          // If deleted, remove FolderId from listened folder Id and renew
          // subscription
          if (deleted && calendarFolderIds.contains(folderEvent.getFolderId())) {
            calendarFolderIds.remove(folderEvent.getFolderId());
          }
        } else {
          if (LOG.isTraceEnabled()) {
            LOG.trace("Folder Event wasn't catched: " + folderEvent.getEventType().name() + "on folder: "
                + folderEvent.getFolderId().getUniqueId());
          }
        }
      }
    }
  }

  private void newSubscription() throws Exception {
    if (LOG.isTraceEnabled()) {
      LOG.trace("New Subscription for user: " + username);
    }
    if (subscription != null) {
      try {
        subscription.unsubscribe();
      } catch (Exception e) {
        // Nothing to do, subscription may be timed out
        if (LOG.isDebugEnabled() || LOG.isTraceEnabled()) {
          LOG.error("Error while unsubscribe, will renew it anyway.", e);
        }
      }
    }
    subscription = integrationService.getService()
                                     .subscribeToPullNotifications(calendarFolderIds,
                                                                   5,
                                                                   null,
                                                                   EventType.Modified,
                                                                   EventType.Created,
                                                                   EventType.Deleted);
  }
}
