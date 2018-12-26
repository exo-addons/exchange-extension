package org.exoplatform.extension.exchange.service;

import java.io.*;
import java.security.KeyStore;
import java.text.ParseException;
import java.util.*;

import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.jcr.Node;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.commons.utils.PropertyManager;
import org.exoplatform.container.PortalContainer;
import org.exoplatform.container.component.ComponentRequestLifecycle;
import org.exoplatform.extension.exchange.service.util.CalendarConverterUtils;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.organization.UserProfile;
import org.exoplatform.web.security.codec.AbstractCodec;
import org.exoplatform.web.security.codec.AbstractCodecBuilder;
import org.exoplatform.web.security.security.TokenServiceInitializationException;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.notification.ItemEvent;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

/**
 * @author Boubaker KHANFIR
 */
@SuppressWarnings("all")
public class IntegrationService {

  private static final String                          GTN_KEY                               = "gtnKey";

  private static final String                          GTN_STORE_PASS                        = "gtnStorePass";

  public static final String                           USER_EXCHANGE_SERVER_URL_ATTRIBUTE    = "exchange.server.url";

  public static final String                           USER_EXCHANGE_SERVER_DOMAIN_ATTRIBUTE = "exchange.server.domain";

  public static final String                           USER_EXCHANGE_USERNAME_ATTRIBUTE      = "exchange.username";

  public static final String                           USER_EXCHANGE_PASSWORD_ATTRIBUTE      = "exchange.password";

  private final static Log                             LOG                                   =
                                                           ExoLogger.getLogger(IntegrationService.class);

  private static final String                          USER_EXCHANGE_HANDLED_ATTRIBUTE       = "exchange.check.date";

  private static final String                          USER_EXO_HANDLED_ATTRIBUTE            = "exo.check.date";

  private static final Map<String, IntegrationService> instances                             =
                                                                 new HashMap<String, IntegrationService>();

  private static AbstractCodec                         codec;

  private final String                                 username;

  private final ExchangeService                        service;

  private final ExoStorageService                      exoStorageService;

  private final ExchangeStorageService                 exchangeStorageService;

  private final CorrespondenceService                  correspondenceService;

  private final OrganizationService                    organizationService;

  private final CalendarService                        calendarService;

  private boolean                                      synchIsCurrentlyRunning               = false;

  public IntegrationService(OrganizationService organizationService,
                            CalendarService calendarService,
                            ExoStorageService exoStorageService,
                            ExchangeStorageService exchangeStorageService,
                            CorrespondenceService correspondenceService,
                            ExchangeService service,
                            String username) {
    this.organizationService = organizationService;
    this.calendarService = calendarService;
    this.exoStorageService = exoStorageService;
    this.exchangeStorageService = exchangeStorageService;
    this.correspondenceService = correspondenceService;
    this.service = service;
    this.username = username;

    // Set corresponding service to each username.
    instances.put(username, this);
  }

  /**
   * Gets user exchange instance service.
   * 
   * @param username
   * @return
   */
  public static IntegrationService getInstance(String username) {
    return instances.get(username);
  }

  /**
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(FolderId folderId) throws Exception {
    return exchangeStorageService.getExchangeCalendar(service, folderId);
  }

  /**
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * 
   * @param folderId
   * @throws Exception
   * @return List of event IDs
   */
  public List<String> synchronizeFullCalendar(FolderId folderId) throws Exception {
    List<String> updatedExoEventIds = new ArrayList<String>();
    CalendarFolder folder = exchangeStorageService.getExchangeCalendar(service, folderId);

    // Create Calendar if not present
    exoStorageService.getOrCreateUserCalendar(username, folder);

    Iterable<Item> items = searchAllItems(folderId);
    synchronizeAllExchangeAppointments(updatedExoEventIds, items);
    deleteExoEventsOutOfSynchronization(folderId);

    return updatedExoEventIds;
  }

  /**
   * Synchronize Exchange Calendar identified by 'folderId' with eXo Calendar.
   * The check is done for events modified since 'lastSyncDate'.
   * 
   * @param folderId
   * @param lastSyncDate
   * @param updatedExoEventIDs
   * @throws Exception
   */
  public void synchronizeModificationsOfCalendar(FolderId folderId,
                                                 Date lastSyncDate,
                                                 List<String> updatedExoEventIDs) throws Exception {
    // Serach modified eXo Calendar events since this date, this is used to
    // force synchronization
    Date exoLastSyncDate = getUserExoLastCheckDate();
    if (exoLastSyncDate == null || exoLastSyncDate.before(lastSyncDate)) {
      exoLastSyncDate = lastSyncDate;
    }

    synchronizeAppointmentsByModificationDate(folderId, lastSyncDate, updatedExoEventIDs);
    synchronizeNewExoEvents(folderId, updatedExoEventIDs, exoLastSyncDate);
    synchronizeExoEventsByModificationDate(folderId, updatedExoEventIDs, exoLastSyncDate);
  }

  /**
   * Gets list of personnal Exchange Calendars.
   * 
   * @return list of FolderId
   * @throws Exception
   */
  public List<FolderId> getAllExchangeCalendars() throws Exception {
    return exchangeStorageService.getAllExchangeCalendars(service);
  }

  /**
   * Checks if eXo associated Calendar is present.
   * 
   * @param usename
   * @return true if present.
   * @throws Exception
   */
  public boolean isCalendarSynchronizedWithExchange(String id) throws Exception {
    return correspondenceService.getCorrespondingId(username, id) != null;
  }

  /**
   * Checks if eXo associated Calendar is present.
   * 
   * @param usename
   * @return true if present.
   * @throws Exception
   */
  public boolean isCalendarPresentInExo(FolderId folderId) throws Exception {
    return exoStorageService.getUserCalendar(username, folderId.getUniqueId()) != null;
  }

  /**
   * Creates or updates or deletes eXo Calendar Event associated to Item, switch
   * state in Exchange.
   * 
   * @param itemEvent
   * @param lastSyncDate
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> createOrUpdateOrDelete(ItemEvent itemEvent, Date lastSyncDate) throws Exception {
    List<CalendarEvent> updatedEvents = null;

    Item item = exchangeStorageService.getItem(service, itemEvent.getItemId());
    if (item == null) {
      exoStorageService.deleteEventByAppointmentID(itemEvent.getItemId().getUniqueId(), username);
    } else if (item instanceof Appointment) {
      Appointment appointment = (Appointment) item;
      String eventId = correspondenceService.getCorrespondingId(username, appointment.getId().getUniqueId());
      long diff = lastSyncDate.getTime() - item.getLastModifiedTime().getTime();
      if (eventId == null) {
        updatedEvents = exoStorageService.createEvent(appointment, username);
      } else {
        updatedEvents = exoStorageService.updateEvent(appointment, username);
      }
    }

    return updatedEvents;
  }

  /**
   * Get corresponding User Calenar from Exchange Folder Id
   * 
   * @param folderId
   * @return
   * @throws Exception
   */
  public Calendar getUserCalendarByExchangeFolderId(FolderId folderId) throws Exception {
    return exoStorageService.getUserCalendar(username, folderId.getUniqueId());
  }

  /**
   * @param eventId
   * @throws Exception
   */
  public void updateOrCreateExchangeCalendarEvent(Node eventNode) throws Exception {
    CalendarEvent event = exoStorageService.getExoEventByNode(eventNode);
    if (isCalendarSynchronizedWithExchange(event.getCalendarId())) {
      updateOrCreateExchangeCalendarEvent(event);
    }
  }

  /**
   * @param eventId
   * @throws Exception
   */
  public boolean updateOrCreateExchangeCalendarEvent(String eventId) throws Exception {
    CalendarEvent event = ((CalendarServiceImpl) calendarService).getDataStorage().getEvent(username, eventId);
    return updateOrCreateExchangeCalendarEvent(event);
  }

  /**
   * @param event
   * @throws Exception
   */
  public boolean updateOrCreateExchangeCalendarEvent(CalendarEvent event) throws Exception {
    String exoMasterId = null;
    if (event.getIsExceptionOccurrence() != null && event.getIsExceptionOccurrence()) {
      exoMasterId = exoStorageService.getExoEventMasterRecurenceByOriginalUUID(event.getOriginalReference());
      if (exoMasterId == null) {
        LOG.error("No master Id was found for occurence: " + event.getSummary() + " with recurrenceId = "
            + event.getRecurrenceId() + ". The event will not be updated.");
      }
    }
    return exchangeStorageService.updateOrCreateExchangeAppointment(username,
                                                                    service,
                                                                    event,
                                                                    exoMasterId,
                                                                    this::appointmentUpdated);
  }

  /**
   * @param eventId
   * @throws Exception
   */
  public void deleteExchangeCalendarEvent(String eventId, String calendarId) throws Exception {
    exchangeStorageService.deleteAppointmentByExoEventId(username, service, eventId, calendarId);
  }

  /**
   * Handle Exchange Calendar Deletion by deleting associated eXo Calendar.
   * 
   * @param folderId Exchange Calendar folderId
   * @return
   * @throws Exception
   */
  public boolean deleteExoCalendar(FolderId folderId) throws Exception {
    Folder folder = exchangeStorageService.getExchangeCalendar(service, folderId);
    if (folder != null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Folder was found, but event seems saying that it was deleted.");
      }
      return false;
    }
    return exoStorageService.deleteCalendar(username, folderId.getUniqueId());
  }

  /**
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeCalendar(String calendarId) throws Exception {
    exchangeStorageService.deleteExchangeFolderByCalenarId(username, service, calendarId);
  }

  public void removeInstance() {
    if (LOG.isTraceEnabled()) {
      LOG.trace("Stop Exchange Integration Service for user: " + username);
    }
    instances.remove(username);
  }

  public List<String> synchronizeExchangeFolderState(List<FolderId> calendarFolderIds,
                                                     boolean synchronizeAllExchangeFolders,
                                                     boolean deleteExoCalendarOnUnsync) throws Exception {
    Iterator<FolderId> iterator = calendarFolderIds.iterator();
    while (iterator.hasNext()) {
      FolderId folderId = (FolderId) iterator.next();
      deleteExoCalendarOutOfSync(deleteExoCalendarOnUnsync, iterator, folderId);
    }

    List<String> updatedCalendarEventIds = null;
    // synchronize added Folders
    if (synchronizeAllExchangeFolders) {
      List<FolderId> folderIds = exchangeStorageService.getAllExchangeCalendars(service);
      for (FolderId folderId : folderIds) {
        // Test if not already fully synchronized
        if (!calendarFolderIds.contains(folderId)) {
          // Delete eXo calendar and recreate it
          exoStorageService.deleteCalendar(username, folderId.getUniqueId());

          List<String> tmpUpdatedCalendarEventIds = synchronizeFullCalendar(folderId);
          if (tmpUpdatedCalendarEventIds != null && !tmpUpdatedCalendarEventIds.isEmpty()) {
            if (updatedCalendarEventIds == null) {
              updatedCalendarEventIds = tmpUpdatedCalendarEventIds;
            } else {
              updatedCalendarEventIds.addAll(tmpUpdatedCalendarEventIds);
            }
          }

          calendarFolderIds.add(folderId);
        }
      }
    } else {
      // Checks if synchronized Exchange Folders have been modified
      List<FolderId> synchronizedFolderIds = getSynchronizedExchangeCalendars();
      for (FolderId folderId : synchronizedFolderIds) {
        // Test if not already fully synchronized
        if (!calendarFolderIds.contains(folderId)) {
          // Delete eXo calendar and recreate it
          exoStorageService.deleteCalendar(username, folderId.getUniqueId());

          List<String> tmpUpdatedCalendarEventIds = synchronizeFullCalendar(folderId);
          if (tmpUpdatedCalendarEventIds != null && !tmpUpdatedCalendarEventIds.isEmpty()) {
            if (updatedCalendarEventIds == null) {
              updatedCalendarEventIds = tmpUpdatedCalendarEventIds;
            } else {
              updatedCalendarEventIds.addAll(tmpUpdatedCalendarEventIds);
            }
          }

          calendarFolderIds.add(folderId);
        }
      }
      Iterator<FolderId> folderIdIterator = calendarFolderIds.iterator();
      while (folderIdIterator.hasNext()) {
        FolderId folderId = (FolderId) folderIdIterator.next();
        if (!synchronizedFolderIds.contains(folderId)) {
          folderIdIterator.remove();
        }
      }
    }
    return updatedCalendarEventIds;
  }

  private void deleteExoCalendarOutOfSync(boolean deleteExoCalendarOnUnsync,
                                          Iterator<FolderId> iterator,
                                          FolderId folderId) throws Exception {
    Folder folder = exchangeStorageService.getExchangeCalendar(service, folderId);
    if (folder == null) {
      // Test if the connection is ok, else the exception is thrown because of
      // interrupted connection
      folder =
             exchangeStorageService.getExchangeCalendar(service,
                                                        FolderId.getFolderIdFromWellKnownFolderName(WellKnownFolderName.Calendar));

      if (folder != null) {
        Calendar calendar = exoStorageService.getUserCalendar(username, folderId.getUniqueId());
        if (calendar != null) {
          if (LOG.isTraceEnabled()) {
            LOG.trace("Folder '" + folderId.getUniqueId()
                + "' was deleted from Exchange, stopping synchronization for this folder.");
          }
          if (deleteExoCalendarOnUnsync) {
            exoStorageService.deleteCalendar(username, folderId.getUniqueId());
          } else {
            correspondenceService.deleteCorrespondingId(username, folderId.getUniqueId());
          }
          // Remove FolderId from synchronized folders in Scheduled Job
          iterator.remove();
        }
      }
    }
  }

  public List<FolderId> getSynchronizedExchangeCalendars() throws Exception {
    List<FolderId> folderIds = new ArrayList<FolderId>();
    List<String> folderIdsString = correspondenceService.getSynchronizedExchangeFolderIds(username);
    for (String folderIdString : folderIdsString) {
      folderIds.add(FolderId.getFolderIdFromString(folderIdString));
    }
    return folderIds;
  }

  public void addFolderToSynchronization(String folderIdString) throws Exception {
    String calendarId = CalendarConverterUtils.getCalendarId(folderIdString);
    correspondenceService.setCorrespondingId(username, calendarId, folderIdString);
  }

  public void deleteFolderFromSynchronization(String folderIdString) throws Exception {
    correspondenceService.deleteCorrespondingId(username, folderIdString);
  }

  public synchronized void setSynchronizationStarted() {
    synchIsCurrentlyRunning = true;
  }

  public synchronized void setSynchronizationStopped() {
    synchIsCurrentlyRunning = false;
  }

  public synchronized boolean isSynchronizationStarted() {
    return synchIsCurrentlyRunning;
  }

  private void deleteExoEventsOutOfSynchronization(FolderId folderId) throws Exception {
    List<CalendarEvent> events = exoStorageService.getUserCalendarEvents(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : events) {
      if (correspondenceService.getCorrespondingId(username, calendarEvent.getId()) == null) {
        exoStorageService.deleteEvent(username, calendarEvent);
      } else {
        String itemId = correspondenceService.getCorrespondingId(username, calendarEvent.getId());
        Item item = exchangeStorageService.getItem(service, itemId);
        if (item == null) {
          exoStorageService.deleteEvent(username, calendarEvent);
        }
      }
    }
  }

  private void synchronizeAllExchangeAppointments(List<String> eventIds, Iterable<Item> items) throws Exception,
                                                                                               ServiceLocalException {
    for (Item item : items) {
      if (item instanceof Appointment) {
        List<CalendarEvent> updatedEvents = exoStorageService.createOrUpdateEvent((Appointment) item, username);
        if (updatedEvents != null && !updatedEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : updatedEvents) {
            eventIds.add(calendarEvent.getId());
          }
        }
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
  }

  /**
   * Sets exo and exchange last check full synchronization operation date
   * 
   * @param username
   * @param time
   * @throws Exception
   */
  public void setUserLastCheckDate(long time) throws Exception {
    if (LOG.isTraceEnabled()) {
      LOG.trace("Set last time check of modified eXChange appointments to '{}.{}' for user {}",
                new Date(time),
                time % 1000,
                username);
    }
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      userProfile.setAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE, String.valueOf(time));

      Date savedTime = getUserExoLastCheckDate();
      if (time > savedTime.getTime()) {
        setUserExoLastCheckDate(time);
      }
      organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  /**
   * Gets last check full synchronization operation date
   * 
   * @return
   * @throws Exception
   */
  public Date getUserLastCheckDate() throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      long time =
                userProfile.getAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE) == null ? 0
                                                                                  : Long.valueOf(userProfile.getAttribute(USER_EXCHANGE_HANDLED_ATTRIBUTE));
      Date lastSyncDate = null;
      if (time > 0) {
        lastSyncDate = new Date(time);
      }
      return lastSyncDate;
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  /**
   * Sets exo last check operation date.
   * 
   * @param username
   * @param time
   * @throws Exception
   */
  public void setUserExoLastCheckDate(long time) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }

    if (LOG.isTraceEnabled()) {
      LOG.trace("Set last time check for modified exo events to '{}.{}' for user {}", new Date(time), time % 1000, username);
    }
    try {
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      long savedTime =
                     userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE) == null ? 0
                                                                                  : Long.valueOf(userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE));
      if (savedTime > 0) {
        userProfile.setAttribute(USER_EXO_HANDLED_ATTRIBUTE, String.valueOf(time));
        organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
      } else if (LOG.isTraceEnabled()) {
        LOG.trace("User '" + username
            + "' exo last check time was not set before, may be the synhronization was not run before or an error occured in the meantime.");
      }

    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  /**
   * Gets exo last check operation date
   * 
   * @return
   * @throws Exception
   */
  public Date getUserExoLastCheckDate() throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      long time =
                userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE) == null ? 0
                                                                             : Long.valueOf(userProfile.getAttribute(USER_EXO_HANDLED_ATTRIBUTE));
      Date lastSyncDate = null;
      if (time > 0) {
        lastSyncDate = new Date(time);
      }
      return lastSyncDate;
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  /**
   * @param organizationService
   * @param username
   * @param name
   * @param value
   * @throws Exception
   */
  public static void setUserArrtibute(OrganizationService organizationService,
                                      String username,
                                      String name,
                                      String value) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      if (USER_EXCHANGE_PASSWORD_ATTRIBUTE.equals(name)) {
        value = encodePassword(value);
      }
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      if (userProfile == null) {
        userProfile = organizationService.getUserProfileHandler().createUserProfileInstance(username);
        organizationService.getUserProfileHandler().saveUserProfile(userProfile, true);
      }
      userProfile.setAttribute(name, value);
      organizationService.getUserProfileHandler().saveUserProfile(userProfile, false);
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  /**
   * @param organizationService
   * @param username
   * @param name
   * @return
   * @throws Exception
   */
  public static String getUserArrtibute(OrganizationService organizationService, String username, String name) throws Exception {
    if (organizationService instanceof ComponentRequestLifecycle) {
      ((ComponentRequestLifecycle) organizationService).startRequest(PortalContainer.getInstance());
    }
    try {
      UserProfile userProfile = organizationService.getUserProfileHandler().findUserProfileByName(username);
      String value = null;
      if (userProfile != null) {
        value = userProfile.getAttribute(name);
        if (value != null && USER_EXCHANGE_PASSWORD_ATTRIBUTE.equals(name)) {
          value = decodePassword(value);
        }
      }
      return value;
    } finally {
      if (organizationService instanceof ComponentRequestLifecycle) {
        ((ComponentRequestLifecycle) organizationService).endRequest(PortalContainer.getInstance());
      }
    }
  }

  public ExchangeService getService() {
    return service;
  }

  private void synchronizeExoEventsByModificationDate(FolderId folderId,
                                                      List<String> updatedExoEventIDs,
                                                      Date exoLastSyncDate) throws Exception {
    List<CalendarEvent> modifiedCalendarEvents = searchCalendarEventsModifiedSince(getUserCalendarByExchangeFolderId(folderId),
                                                                                   exoLastSyncDate);
    if (LOG.isTraceEnabled() && modifiedCalendarEvents.size() > 0) {
      LOG.trace("Check exo user calendar for user '{}' since '{}', items found: {}",
                username,
                exoLastSyncDate,
                modifiedCalendarEvents.size());
    }
    for (CalendarEvent calendarEvent : modifiedCalendarEvents) {
      // If modified with synchronization, ignore
      if (updatedExoEventIDs.contains(calendarEvent.getId())) {
        continue;
      }
      String exoMasterId = null;
      if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
        exoMasterId = exoStorageService.getExoEventMasterRecurenceByOriginalUUID(calendarEvent.getOriginalReference());
        if (exoMasterId == null) {
          LOG.error("No master Id was found for occurence: " + calendarEvent.getSummary() + " with recurrenceId = "
              + calendarEvent.getRecurrenceId() + ". The event will not be updated.");
        }
      }
      boolean deleteEvent = exchangeStorageService.updateOrCreateExchangeAppointment(username,
                                                                                     service,
                                                                                     calendarEvent,
                                                                                     exoMasterId,
                                                                                     this::appointmentUpdated);
      if (deleteEvent) {
        exoStorageService.deleteEvent(username, calendarEvent);
      }
      updatedExoEventIDs.add(calendarEvent.getId());
    }
  }

  @SuppressWarnings("deprecation")
  private void synchronizeNewExoEvents(FolderId folderId,
                                       List<String> updatedExoEventIDs,
                                       Date exoLastSyncDate) throws Exception {
    // Search for existant Appointments in Exchange but not in eXo
    Iterable<CalendarEvent> unsynchronizedEvents = searchUnsynchronizedAppointments(username, folderId.getUniqueId());
    for (CalendarEvent calendarEvent : unsynchronizedEvents) {
      // To not have redendance
      if (updatedExoEventIDs != null && updatedExoEventIDs.contains(calendarEvent.getId())) {
        continue;
      }

      if (calendarEvent.getLastUpdatedTime() != null) {
        if (calendarEvent.getLastUpdatedTime().after(exoLastSyncDate)) {
          String exoMasterId = null;
          if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
            exoMasterId = exoStorageService.getExoEventMasterRecurenceByOriginalUUID(calendarEvent.getOriginalReference());
            if (exoMasterId == null) {
              LOG.error("No master Id was found for occurence: " + calendarEvent.getSummary() + " with recurrenceId = "
                  + calendarEvent.getRecurrenceId() + ". The event will not be updated.");
              continue;
            }
          }
          boolean deleteEvent = exchangeStorageService.updateOrCreateExchangeAppointment(username,
                                                                                         service,
                                                                                         calendarEvent,
                                                                                         exoMasterId,
                                                                                         this::appointmentUpdated);
          if (deleteEvent) {
            exoStorageService.deleteEvent(username, calendarEvent);
          }
          if (updatedExoEventIDs != null) {
            updatedExoEventIDs.add(calendarEvent.getId());
          }
        } else {
          exoStorageService.deleteEvent(username, calendarEvent);
        }
      }
    }
  }

  private void synchronizeAppointmentsByModificationDate(FolderId folderId,
                                                         Date lastSyncDate,
                                                         List<String> updatedExoEventIDs) throws Exception,
                                                                                          ServiceLocalException,
                                                                                          ParseException {
    Iterable<Item> items = searchAllAppointmentsModifiedSince(folderId, lastSyncDate);
    // Search for modified Appointments in Exchange, since last check date.
    for (Item item : items) {
      if (item instanceof Appointment) {
        if (lastSyncDate != null && item.getLastModifiedTime() != null && item.getLastModifiedTime().before(lastSyncDate)) {
          continue;
        }
        // Test if there is a modification conflict
        CalendarEvent event = exoStorageService.getEventByAppointmentId(username, item.getId().getUniqueId());
        if (event != null) {
          if (updatedExoEventIDs != null && updatedExoEventIDs.contains(event.getId())) {
            // Already updated by previous operation
            continue;
          }
          @SuppressWarnings("deprecation")
          long eventModifDate = event.getLastModified();
          long itemModifDate = item.getLastModifiedTime().getTime();
          if (itemModifDate > eventModifDate) {
            List<CalendarEvent> updatedEvents = exoStorageService.updateEvent((Appointment) item, username);
            if (updatedEvents != null && !updatedEvents.isEmpty() && updatedExoEventIDs != null) {
              for (CalendarEvent calendarEvent : updatedEvents) {
                updatedExoEventIDs.add(calendarEvent.getId());
              }
            }
          }
        } else {
          List<CalendarEvent> updatedEvents = exoStorageService.createEvent((Appointment) item, username);
          if (updatedEvents != null && !updatedEvents.isEmpty() && updatedExoEventIDs != null) {
            for (CalendarEvent calendarEvent : updatedEvents) {
              updatedExoEventIDs.add(calendarEvent.getId());
            }
          }
        }
      } else {
        LOG.warn("Item bound from exchange but not of type 'Appointment':" + item.getItemClass());
      }
    }
  }

  private Iterable<CalendarEvent> searchUnsynchronizedAppointments(String username, String folderId) throws Exception {
    List<CalendarEvent> calendarEvents = exoStorageService.getUserCalendarEvents(username, folderId);
    Iterator<CalendarEvent> calendarEventsIterator = calendarEvents.iterator();
    while (calendarEventsIterator.hasNext()) {
      CalendarEvent calendarEvent = calendarEventsIterator.next();
      String itemId = correspondenceService.getCorrespondingId(username, calendarEvent.getId());
      if (itemId == null) {
        // Item was detected, and will be created
        continue;
      }
      Item item = exchangeStorageService.getItem(service, ItemId.getItemIdFromString(itemId));
      if (item != null) {
        calendarEventsIterator.remove();
      }
    }
    return calendarEvents;
  }

  private List<Item> searchAllItems(FolderId parentFolderId) throws Exception {
    ItemView view = new ItemView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId, view);
    int totalCount = findResults.getTotalCount();
    if (LOG.isTraceEnabled() && totalCount > 0) {
      LOG.trace("Check all exchange calendar for user '{}', items found: {}", username, totalCount);
    }
    return findResults.getItems();
  }

  private List<CalendarEvent> searchCalendarEventsModifiedSince(Calendar calendar, Date date) throws Exception {
    if (date == null) {
      return exoStorageService.getAllExoEvents(username, calendar);
    }
    return exoStorageService.findExoEventsModifiedSince(username, calendar, date);
  }

  private List<Item> searchAllAppointmentsModifiedSince(FolderId parentFolderId, Date date) throws Exception {
    if (date == null) {
      return searchAllItems(parentFolderId);
    }

    java.util.Calendar calendar = java.util.Calendar.getInstance();
    calendar.setTime(date);

    ItemView view = new ItemView(100);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindItemsResults<Item> findResults = service.findItems(parentFolderId,
                                                           new SearchFilter.IsGreaterThan(ItemSchema.LastModifiedTime,
                                                                                          calendar.getTime()),
                                                           view);
    int totalCount = findResults.getTotalCount();
    if (LOG.isTraceEnabled() && totalCount > 0) {
      LOG.trace("Check exchange user calendar for user '{}' since '{}', items found: {}", username, date, totalCount);
    }
    return findResults.getItems();
  }

  private static void initCodec() throws TokenServiceInitializationException {
    String builderType = PropertyManager.getProperty("gatein.codec.builderclass");
    Map<String, String> config = new HashMap<>();

    if (builderType != null) {
      // If there is config for codec in configuration.properties, we read the
      // config parameters from config file
      // referenced in configuration.properties
      String configFilePath = PropertyManager.getProperty("gatein.codec.config");
      File configFile = new File(configFilePath);
      try (InputStream in = new FileInputStream(configFile)) {
        Properties properties = new Properties();
        properties.load(in);
        for (Map.Entry<?, ?> entry : properties.entrySet()) {
          config.put((String) entry.getKey(), (String) entry.getValue());
        }
        config.put("gatein.codec.config.basedir", configFile.getParentFile().getAbsolutePath());
      } catch (IOException e) {
        throw new TokenServiceInitializationException("Failed to read the config parameters from file '" + configFilePath + "'.",
                                                      e);
      }
    } else {
      // If there is no config for codec in configuration.properties, we
      // generate key if it does not exist and setup the
      // default config
      builderType = "org.exoplatform.web.security.codec.JCASymmetricCodecBuilder";
      String gtnConfDir = PropertyManager.getProperty("gatein.conf.dir");
      if (gtnConfDir == null || gtnConfDir.length() == 0) {
        throw new TokenServiceInitializationException("'gatein.conf.dir' property must be set.");
      }
      File f = new File(gtnConfDir + "/codec/codeckey.txt");
      if (!f.exists()) {
        File codecDir = f.getParentFile();
        if (!codecDir.exists()) {
          codecDir.mkdir();
        }
        try (OutputStream out = new FileOutputStream(f)) {
          KeyGenerator keyGen = KeyGenerator.getInstance("AES");
          keyGen.init(128);
          SecretKey key = keyGen.generateKey();
          KeyStore store = KeyStore.getInstance("JCEKS");
          store.load(null, GTN_STORE_PASS.toCharArray());
          store.setEntry(GTN_KEY, new KeyStore.SecretKeyEntry(key), new KeyStore.PasswordProtection("gtnKeyPass".toCharArray()));
          store.store(out, GTN_STORE_PASS.toCharArray());
        } catch (Exception e) {
          throw new TokenServiceInitializationException(e);
        }
      }
      config.put("gatein.codec.jca.symmetric.keyalg", "AES");
      config.put("gatein.codec.jca.symmetric.keystore", "codeckey.txt");
      config.put("gatein.codec.jca.symmetric.storetype", "JCEKS");
      config.put("gatein.codec.jca.symmetric.alias", GTN_KEY);
      config.put("gatein.codec.jca.symmetric.keypass", "gtnKeyPass");
      config.put("gatein.codec.jca.symmetric.storepass", GTN_STORE_PASS);
      config.put("gatein.codec.config.basedir", f.getParentFile().getAbsolutePath());
    }

    try {
      codec = Class.forName(builderType).asSubclass(AbstractCodecBuilder.class).newInstance().build(config);
    } catch (Exception e) {
      throw new TokenServiceInitializationException("Could not initialize CookieTokenService.codec.", e);
    }
  }

  private static String decodePassword(String password) throws TokenServiceInitializationException {
    if (codec == null) {
      initCodec();
    }
    return codec.decode(password);
  }

  private static String encodePassword(String password) throws TokenServiceInitializationException {
    if (codec == null) {
      initCodec();
    }
    return codec.encode(password);
  }

  /**
   * Make sure that eXo Event has the same modification date than Exchange event
   * 
   * @param appointment
   * @return
   */
  public Boolean appointmentUpdated(Appointment appointment) {
    try {
      // Refresh Appointment
      appointment = (Appointment) exchangeStorageService.getItem(service, appointment.getId());

      CalendarEvent event = exoStorageService.getEventByAppointmentId(username, appointment.getId().getUniqueId());
      if (event != null) {
        exoStorageService.updateModifiedDateOfEvent(username, event, appointment.getLastModifiedTime());
      }
    } catch (Exception e) {
      LOG.warn("Error occurred while updating eXo Event last updated date", e);
      return false;
    }
    return true;
  }
}
