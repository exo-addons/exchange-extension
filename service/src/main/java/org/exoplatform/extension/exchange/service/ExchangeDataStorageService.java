package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.*;
import java.util.Map.Entry;
import java.util.function.Function;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;

import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.commons.utils.CommonsUtils;
import org.exoplatform.container.PortalContainer;
import org.exoplatform.extension.exchange.service.util.CalendarConverterUtils;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;

import microsoft.exchange.webservices.data.core.*;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.*;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceObjectPropertyException;
import microsoft.exchange.webservices.data.core.exception.service.remote.ServiceResponseException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.property.definition.PropertyDefinition;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FolderView;

/**
 * @author Boubaker Khanfir
 */
@SuppressWarnings("all")
public class ExchangeDataStorageService implements Serializable {
  private static final long     serialVersionUID = 6348129698208975430L;

  private static final Log      LOG              = ExoLogger.getLogger(ExchangeDataStorageService.class);

  private OrganizationService   organizationService;

  private CorrespondenceService correspondenceService;

  public ExchangeDataStorageService(CorrespondenceService correspondenceService) {
    this.correspondenceService = correspondenceService;
  }

  /**
   * Gets list of personnal Exchange Calendars.
   * 
   * @return list of FolderId
   * @throws Exception
   */
  public List<FolderId> getAllExchangeCalendars(ExchangeService service) throws Exception {
    List<FolderId> calendarFolderIds = new ArrayList<FolderId>();
    CalendarFolder calendarRootFolder = CalendarFolder.bind(service, WellKnownFolderName.Calendar);

    calendarFolderIds.add(calendarRootFolder.getId());
    List<Folder> calendarfolders = searchSubFolders(service, calendarRootFolder.getId());

    if (calendarfolders != null && !calendarfolders.isEmpty()) {
      for (Folder tmpFolder : calendarfolders) {
        calendarFolderIds.add(tmpFolder.getId());
      }
    }
    return calendarFolderIds;
  }

  /**
   * @param username
   * @param service
   * @param event
   * @param exoMasterId
   * @return true if the CalenarEvent have to be deleted
   * @throws Exception
   */
  public boolean updateOrCreateExchangeAppointment(String username,
                                                   ExchangeService service,
                                                   CalendarEvent event,
                                                   String exoMasterId,
                                                   Function<Appointment, Boolean> appointmentSavedCallback) throws Exception {

    if (event == null) {
      return false;
    }
    String folderIdString = correspondenceService.getCorrespondingId(username, event.getCalendarId());
    if (folderIdString == null || folderIdString.isEmpty()) {
      LOG.trace("eXo Calendar with id '{}' is not synhronized with Exchange, ignore Event: {}",
                event.getCalendarId(),
                event.getSummary());
      return false;
    }

    String itemId = correspondenceService.getCorrespondingId(username, event.getId());
    boolean isNew = true;
    Appointment appointment = null;
    if (itemId != null) {
      try {
        appointment = Appointment.bind(service, ItemId.getItemIdFromString(itemId));
        isNew = appointment == null;
      } catch (ServiceResponseException e) {
        LOG.warn("Item was not bound, it was deleted or not yet created:" + event.getId(), e);
        correspondenceService.deleteCorrespondingId(username, event.getId());
      }
    }

    if (!isNew && appointment.getLastModifiedTime() != null
        && appointment.getLastModifiedTime().getTime() == event.getLastModified()) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("IGNORE updating appointment '{}' because its modified date is same as event modified date",
                  event.getSummary());
      }
      return false;
    }

    if (event.getRecurrenceId() == null
        && (event.getRepeatType() == null || event.getRepeatType().equals(CalendarEvent.RP_NOREPEAT))) {
      if (isNew) {
        // Checks if this event was already in Exchange, if it's the case, it
        // means that the item was not found because the user has removed it
        // from
        // Exchange
        if (CalendarConverterUtils.isExchangeEventId(event.getId())) {
          LOG.error("Conflict in modification, inconsistant data, the event was deleted in Exchange but seems always in eXo, the event will be deleted from Exchange.");
          deleteAppointmentByExoEventId(username, service, event.getId(), event.getCalendarId());
          return false;
        }
        appointment = new Appointment(service);
      }

      CalendarConverterUtils.convertExoToExchangeEvent(appointment, event, username, getOrganizationService().getUserHandler());
    } else {
      if ((event.getRecurrenceId() != null && !event.getRecurrenceId().isEmpty())
          || (event.getIsExceptionOccurrence() != null && event.getIsExceptionOccurrence())) {
        if (isNew) {
          String exchangeMasterId = correspondenceService.getCorrespondingId(username, exoMasterId);
          if (StringUtils.isNotBlank(exchangeMasterId)) {
            appointment = getAppointmentOccurence(service, exchangeMasterId, event.getRecurrenceId());
            if (appointment != null) {
              correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
              isNew = false;
            }
          }
          if (isNew) {
            appointment = new Appointment(service);
          }
        }

        CalendarConverterUtils.convertExoToExchangeOccurenceEvent(appointment,
                                                                  event,
                                                                  username,
                                                                  getOrganizationService().getUserHandler());
      } else {
        if (isNew) {
          // Checks if this event was already in Exchange, if it's the case, it
          // means that the item was not found because the user has removed it
          // from Exchange
          if (CalendarConverterUtils.isExchangeEventId(event.getId())) {
            LOG.error("Conflict in modification, inconsistant data, the event was deleted in Exchange but seems always in eXo, the event will be deleted from Exchange.");
            deleteAppointmentByExoEventId(username, service, event.getId(), event.getCalendarId());
            return false;
          }
          appointment = new Appointment(service);
        }
        List<Appointment> toDeleteOccurences =
                                             CalendarConverterUtils.convertExoToExchangeMasterRecurringCalendarEvent(appointment,
                                                                                                                     event,
                                                                                                                     username,
                                                                                                                     getOrganizationService().getUserHandler());

        if (toDeleteOccurences != null && !toDeleteOccurences.isEmpty()) {
          for (Appointment occAppointment : toDeleteOccurences) {

            CalendarServiceImpl calendarService =
                                                (CalendarServiceImpl) PortalContainer.getInstance()
                                                                                     .getComponentInstanceOfType(CalendarService.class);
            CalendarEvent tmpEvent = CalendarConverterUtils.getOccurenceOfDate(username,
                                                                               calendarService.getDataStorage(),
                                                                               event,
                                                                               occAppointment.getOriginalStart());
            if (tmpEvent != null) {
              CalendarConverterUtils.convertExoToExchangeOccurenceEvent(occAppointment,
                                                                        tmpEvent,
                                                                        username,
                                                                        getOrganizationService().getUserHandler());
              if (LOG.isDebugEnabled()) {
                LOG.debug("CREATE for user {} an Exchange Exceptional Occurence Appointment: {}. From: {} To: {} AlldayEvent = {}",
                          username,
                          tmpEvent.getSummary(),
                          getStartDate(occAppointment),
                          getEndDate(occAppointment),
                          getIsAllDayEvent(occAppointment));
              }
              try {
                occAppointment.update(ConflictResolutionMode.AutoResolve);
                displayLogIfNotSameDates(occAppointment);
              } catch (ServiceResponseException e) {
                if (e.getMessage() != null && e.getMessage().contains("At least one recipient isn't valid")) {
                  if (LOG.isTraceEnabled()) {
                    LOG.warn("Error while saving appointment", e);
                  }
                  occAppointment.update(ConflictResolutionMode.AutoResolve, SendInvitationsOrCancellationsMode.SendToNone);
                  displayLogIfNotSameDates(occAppointment);
                } else {
                  throw e;
                }
              }
              correspondenceService.setCorrespondingId(username, tmpEvent.getId(), occAppointment.getId().getUniqueId());
              appointmentSavedCallback.apply(occAppointment);
              continue;
            }
            // Verify if deleted occurences is an exception existing occurence
            // or not
            String exoId = correspondenceService.getCorrespondingId(username, occAppointment.getId().getUniqueId());
            if (exoId == null) {
              deleteAppointment(username, service, occAppointment.getId());
            }
          }
        }
      }
    }
    if (isNew) {
      if (LOG.isDebugEnabled()) {
        LOG.debug("CREATE for user '{}' Exchange Appointment: {}. From: {} To: {} AlldayEvent = {}",
                  username,
                  event.getSummary(),
                  getStartDate(appointment),
                  getEndDate(appointment),
                  getIsAllDayEvent(appointment));
      }
      FolderId folderId = FolderId.getFolderIdFromString(folderIdString);
      try {
        appointment.save(folderId);
        displayLogIfNotSameDates(appointment);
      } catch (ServiceResponseException e) {
        if (e.getMessage() != null && e.getMessage().contains("At least one recipient isn't valid")) {
          if (LOG.isTraceEnabled()) {
            LOG.warn("Error while saving appointment", e);
          }
          appointment.save(folderId, SendInvitationsMode.SendToNone);
        } else {
          throw e;
        }
      }
      correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
      appointmentSavedCallback.apply(appointment);
    } else {
      if (getLastModifiedDate(appointment).getTime() == event.getLastModified()) {
        if (LOG.isDebugEnabled()) {
          LOG.debug("IGNORE UPDATE for user '{}' Exchange Appointment '{}' because its modified date is the same as eXo Event. From: {} To: {} AlldayEvent = {}",
                    username,
                    event.getSummary(),
                    getStartDate(appointment),
                    getEndDate(appointment),
                    getIsAllDayEvent(appointment));
        }
        return false;
      } else if (LOG.isDebugEnabled()) {
        LOG.debug("UPDATE for user '{}' Exchange Appointment: {}. From: {} To: {} AlldayEvent = {}",
                  username,
                  event.getSummary(),
                  getStartDate(appointment),
                  getEndDate(appointment),
                  getIsAllDayEvent(appointment));
      }
      try {
        appointment.update(ConflictResolutionMode.AutoResolve);
        displayLogIfNotSameDates(appointment);
      } catch (ServiceResponseException e) {
        if (e.getMessage() != null && e.getMessage().contains("At least one recipient isn't valid")) {
          if (LOG.isTraceEnabled()) {
            LOG.warn("Error while saving appointment", e);
          }
          appointment.update(ConflictResolutionMode.AutoResolve, SendInvitationsOrCancellationsMode.SendToNone);
          displayLogIfNotSameDates(appointment);
        } else {
          throw e;
        }
      }
      correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
      appointmentSavedCallback.apply(appointment);
    }
    return false;
  }

  /**
   * @param username
   * @param service
   * @param eventId
   * @param calendarId
   * @throws Exception
   */
  public void deleteAppointmentByExoEventId(String username,
                                            ExchangeService service,
                                            String eventId,
                                            String calendarId) throws Exception {
    String itemId = correspondenceService.getCorrespondingId(username, eventId);
    if (itemId == null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("The event was deleted from eXo but seems don't have corresponding Event in Exchange, ignore.");
      }
    } else {
      // Verify that calendar is synchronized
      if (correspondenceService.getCorrespondingId(username, calendarId) == null) {
        LOG.warn("Calendar with id '" + calendarId + "' seems not synchronized with exchange.");

        // delete correspondance for safe operations while sync is activated
        correspondenceService.deleteCorrespondingId(username, eventId);
      } else {
        deleteAppointment(username, service, itemId);
      }
    }
  }

  /**
   * @param username
   * @param service
   * @param itemId
   * @throws Exception
   */
  public void deleteAppointment(String username, ExchangeService service, String itemId) throws Exception {
    deleteAppointment(username, service, ItemId.getItemIdFromString(itemId));
  }

  /**
   * @param username
   * @param service
   * @param itemId
   * @throws Exception
   */
  public void deleteAppointment(String username, ExchangeService service, ItemId itemId) throws Exception {
    Appointment appointment = null;
    try {
      appointment = Appointment.bind(service, itemId);
      if (LOG.isDebugEnabled()) {
        LOG.debug("DELETE for user '{}' Exchange appointment: {}", username, appointment.getSubject());
      }
      appointment.delete(DeleteMode.HardDelete);
    } catch (ServiceResponseException e) {
      if (LOG.isDebugEnabled()) {
        LOG.debug("Exchange Item was not bound for user '{}', it was deleted or not yet created: {}", username, itemId);
      }
    }
    correspondenceService.deleteCorrespondingId(username, itemId.getUniqueId());
  }

  /**
   * @param username
   * @param service
   * @param calendarId
   * @throws Exception
   */
  public void deleteExchangeFolderByCalenarId(String username, ExchangeService service, String calendarId) throws Exception {
    if (CalendarConverterUtils.isExchangeCalendarId(calendarId)) {
      LOG.warn("Can't delete Exchange Calendar, because it was created on Exchange: " + calendarId);
      return;
    }
    String folderId = correspondenceService.getCorrespondingId(username, calendarId);
    if (folderId == null) {
      LOG.warn("Conflict in modification, inconsistant data, the Calendar was deleted from eXo but seems don't have corresponding Folder in Exchange, ignore.");
      return;
    } else {
      Folder folder = null;
      try {
        folder = Folder.bind(service, FolderId.getFolderIdFromString(folderId));
        LOG.trace("DELETE Exchange folder: " + folder.getDisplayName());
        folder.delete(DeleteMode.MoveToDeletedItems);
      } catch (ServiceResponseException e) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("Exchange Folder was not bound, it was deleted or not yet created:" + folderId);
        }
      }
      correspondenceService.deleteCorrespondingId(username, calendarId);
    }
  }

  /**
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(ExchangeService service, String folderId) throws Exception {
    return getExchangeCalendar(service, FolderId.getFolderIdFromString(folderId));
  }

  /**
   * @param folderId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public CalendarFolder getExchangeCalendar(ExchangeService service, FolderId folderId) throws Exception {
    CalendarFolder folder = null;
    try {
      folder = CalendarFolder.bind(service, folderId);
    } catch (ServiceResponseException e) {
      LOG.warn("Can't get Folder identified by id: " + folderId.getUniqueId());
    }
    return folder;
  }

  /**
   * @param service
   * @param itemId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public Item getItem(ExchangeService service, String itemId) throws Exception {
    return getItem(service, ItemId.getItemIdFromString(itemId));
  }

  /**
   * @param service
   * @param itemId
   * @return Exchange Folder instance based on Exchange FolderId object
   * @throws Exception
   */
  public Item getItem(ExchangeService service, ItemId itemId) throws Exception {
    Item item = null;
    try {
      item = Item.bind(service, itemId);
    } catch (ServiceResponseException e) {
      LOG.debug("Can't get item for user '{}' identified by id: {}", itemId.getUniqueId());
    }
    return item;
  }

  private List<Folder> searchSubFolders(ExchangeService service, FolderId parentFolderId) throws Exception {
    FolderView view = new FolderView(1000);
    view.setPropertySet(new PropertySet(BasePropertySet.FirstClassProperties));
    FindFoldersResults findResults = service.findFolders(parentFolderId, view);
    return findResults.getFolders();
  }

  private Appointment getAppointmentOccurence(ExchangeService service,
                                              String exchangeMasterId,
                                              String recurrenceId) throws Exception {
    Appointment masterAppointment = Appointment.bind(service,
                                                     ItemId.getItemIdFromString(exchangeMasterId),
                                                     new PropertySet(AppointmentSchema.Recurrence));
    return CalendarConverterUtils.getAppointmentOccurence(masterAppointment, recurrenceId);
  }

  private Date getLastModifiedDate(Appointment appointment) throws Exception {
    try {
      return appointment.getLastModifiedTime();
    } catch (ServiceObjectPropertyException e) {
      Appointment appointmentWithModifiedDate = Appointment.bind(appointment.getService(),
                                                                 appointment.getId(),
                                                                 new PropertySet(AppointmentSchema.LastModifiedTime));
      return appointmentWithModifiedDate.getLastModifiedTime();
    }
  }

  private Date getStartDate(Appointment appointment) throws ServiceLocalException {
    try {
      return appointment.getStart();
    } catch (ServiceLocalException e) {
      return null;
    }
  }

  private Date getEndDate(Appointment appointment) {
    try {
      return appointment.getEnd();
    } catch (ServiceLocalException e) {
      return null;
    }
  }

  private Boolean getIsAllDayEvent(Appointment appointment) {
    try {
      return appointment.getIsAllDayEvent();
    } catch (ServiceLocalException e) {
      return false;
    }
  }

  private void displayLogIfNotSameDates(Appointment appointment) throws Exception,
                                                                 ServiceLocalException,
                                                                 ServiceResponseException {
    // displayAppointmentProperties(appointment);
    try {
      Appointment appointmentUpdated = Appointment.bind(appointment.getService(), appointment.getId());
      if (appointment.getIsAllDayEvent() && (!DateUtils.isSameInstant(getStartDate(appointmentUpdated), getStartDate(appointment))
          || !DateUtils.isSameInstant(getEndDate(appointmentUpdated), getEndDate(appointment)))) {
        LOG.debug("Not same date between updated and to update FROM {}/{} TO {}/{}, update again",
                  getStartDate(appointmentUpdated),
                  getStartDate(appointment),
                  getEndDate(appointmentUpdated),
                  getEndDate(appointment));
        // displayAppointmentProperties(appointmentUpdated);
        appointment.update(ConflictResolutionMode.AutoResolve);
      }
    } catch (Exception e) {
      LOG.debug("Error re-saving appointment");
    }
  }

  private void displayAppointmentProperties(Appointment appointment) {
    PropertyBag properties = appointment.getPropertyBag();
    if (properties != null && properties.getProperties() != null) {
      Set<Entry<PropertyDefinition, Object>> entrySet = properties.getProperties().entrySet();
      for (Entry<PropertyDefinition, Object> entry : entrySet) {
        String uri = entry.getKey().getUri();
        uri = uri != null && uri.indexOf(":") > 0 ? uri.split(":")[1] : uri;
        System.out.println("              Property " + entry.getKey().getUri() + " = " + entry.getValue());
      }
    }
  }

  private OrganizationService getOrganizationService() {
    if (organizationService == null) {
      organizationService = CommonsUtils.getService(OrganizationService.class);
    }
    return organizationService;
  }
}
