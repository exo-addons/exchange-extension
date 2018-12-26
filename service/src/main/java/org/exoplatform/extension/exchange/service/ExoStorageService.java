package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import javax.jcr.Node;
import javax.jcr.NodeIterator;
import javax.jcr.query.*;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.calendar.service.CalendarEvent;
import org.exoplatform.calendar.service.CalendarService;
import org.exoplatform.calendar.service.impl.CalendarServiceImpl;
import org.exoplatform.calendar.service.impl.JCRDataStorage;
import org.exoplatform.calendar.util.Constants;
import org.exoplatform.commons.utils.ISO8601;
import org.exoplatform.extension.exchange.listener.CalendarCreateUpdateAction;
import org.exoplatform.extension.exchange.service.util.CalendarConverterUtils;
import org.exoplatform.services.jcr.ext.common.SessionProvider;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;

import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceObjectPropertyException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;

/**
 * @author Boubaker Khanfir
 */
@SuppressWarnings("all")
public class ExoStorageService implements Serializable {
  private static final String           EXO_DATETIME_PROPERTY         = "exo:datetime";

  private static final long             serialVersionUID              = 6614108102985034995L;

  private static final Log              LOG                           = ExoLogger.getLogger(ExoStorageService.class);

  @SuppressWarnings("all")
  private static final DateFormat       EXCLUDE_ID_FORMAT_FIRST_CHARS = new SimpleDateFormat("yyyyMMdd");

  private transient JCRDataStorage      storage;

  private transient OrganizationService organizationService;

  private CorrespondenceService         correspondenceService;

  public ExoStorageService(OrganizationService organizationService,
                           CalendarService calendarService,
                           CorrespondenceService correspondenceService) {
    this.storage = ((CalendarServiceImpl) calendarService).getDataStorage();
    this.organizationService = organizationService;
    this.correspondenceService = correspondenceService;
  }

  /**
   * Deletes eXo Calendar Event that corresponds to given appointment Id.
   * 
   * @param appointmentId
   * @param username
   * @throws Exception
   */
  public void deleteEventByAppointmentID(String appointmentId, String username) throws Exception {
    @SuppressWarnings("all")
    CalendarEvent calendarEvent = getEventByAppointmentId(username, appointmentId);
    if (calendarEvent != null) {
      deleteEvent(username, calendarEvent);
    }
  }

  /**
   * Deletes eXo Calendar Event.
   * 
   * @param username
   * @param calendarEvent
   * @throws Exception
   */
  @SuppressWarnings("all")
  public void deleteEvent(String username, CalendarEvent calendarEvent) throws Exception {
    if (calendarEvent == null) {
      LOG.warn("Event is null, can't delete it for username: " + username);
      return;
    }

    if ((calendarEvent.getRepeatType() == null || calendarEvent.getRepeatType().equals(CalendarEvent.RP_NOREPEAT))
        && (calendarEvent.getIsExceptionOccurrence() == null || !calendarEvent.getIsExceptionOccurrence())) {
      if (LOG.isTraceEnabled()) {
        LOG.debug("DELETE user calendar event: " + calendarEvent.getSummary());
      }
      storage.removeUserEvent(username, calendarEvent.getCalendarId(), calendarEvent.getId());
      // Remove correspondence between exo and exchange IDs
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    } else if (calendarEvent.getIsExceptionOccurrence() != null && calendarEvent.getIsExceptionOccurrence()) {
      if (LOG.isDebugEnabled()) {
        LOG.debug("DELETE user calendar event exceptional occurence: " + calendarEvent.getSummary() + ", id="
            + calendarEvent.getRecurrenceId());
      }
      storage.removeUserEvent(username, calendarEvent.getCalendarId(), calendarEvent.getId());
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    } else if (calendarEvent.getRecurrenceId() != null && !calendarEvent.getRecurrenceId().isEmpty()) {
      if (LOG.isDebugEnabled()) {
        LOG.debug("DELETE user calendar event occurence from series: " + calendarEvent.getSummary() + " with id : "
            + calendarEvent.getRecurrenceId());
      }
      storage.removeOccurrenceInstance(username, calendarEvent);
    } else {
      if (LOG.isDebugEnabled()) {
        LOG.debug("DELETE user calendar event series: " + calendarEvent.getSummary());
      }
      storage.removeRecurrenceSeries(username, calendarEvent);
      // Remove correspondence between exo and exchange IDs
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    }
  }

  /**
   * Delete eXo Calendar.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  @SuppressWarnings("all")
  public boolean deleteCalendar(String username, String folderId) throws Exception {
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    if (calendarId == null) {
      calendarId = CalendarConverterUtils.getCalendarId(folderId);
    }
    List<CalendarEvent> events = getUserCalendarEvents(username, folderId);
    if (events == null) {
      return false;
    }
    for (CalendarEvent calendarEvent : events) {
      correspondenceService.deleteCorrespondingId(username, calendarEvent.getId());
    }
    storage.removeUserCalendar(username, calendarId);
    correspondenceService.deleteCorrespondingId(username, folderId, calendarId);
    return true;
  }

  /**
   * Gets User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  public Calendar getUserCalendar(String username, String folderId) throws Exception {
    return getUserCalendar(username, folderId, true);
  }

  /**
   * Gets User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @param deleteIfCorrespondentExists
   * @return
   * @throws Exception
   */
  public Calendar getUserCalendar(String username, String folderId, boolean deleteIfCorrespondentExists) throws Exception {
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    Calendar calendar = null;
    if (calendarId != null) {
      calendar = storage.getUserCalendar(username, calendarId);
      if (calendar == null && deleteIfCorrespondentExists) {
        correspondenceService.deleteCorrespondingId(username, folderId);
      }
    }
    return calendar;
  }

  /**
   * Gets User Calendar identified by Exchange folder Id, or creates it if not
   * existing.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  public Calendar getOrCreateUserCalendar(String username, Folder folder) throws Exception {
    Calendar calendar = getUserCalendar(username, folder.getId().getUniqueId(), false);
    String calendarId = CalendarConverterUtils.getCalendarId(folder.getId().getUniqueId());
    if (calendar == null) {
      Calendar tmpCalendar = storage.getUserCalendar(username, calendarId);
      if (tmpCalendar != null) {
        // Renew Calendar
        storage.removeUserCalendar(username, calendarId);
      }
    }
    if (calendar == null) {
      if (LOG.isTraceEnabled()) {
        LOG.debug("CREATE user calendar from Exchange: " + folder.getDisplayName());
      }

      calendar = new Calendar();
      calendar.setId(calendarId);
      calendar.setName(CalendarConverterUtils.getCalendarName(folder.getDisplayName()));
      calendar.setCalendarOwner(username);
      calendar.setDataInit(false);
      calendar.setEditPermission(new String[] { "any read" });
      calendar.setCalendarColor(Constants.COLORS[(int) (Math.random() * Constants.COLORS.length)]);

      storage.saveUserCalendar(username, calendar, true);

      // Set IDs correspondence
      correspondenceService.setCorrespondingId(username, calendar.getId(), folder.getId().getUniqueId());
    }
    return calendar;
  }

  /**
   * Gets Events from User Calendar identified by Exchange folder Id.
   * 
   * @param username
   * @param folderId
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> getUserCalendarEvents(String username, String folderId) throws Exception {
    List<CalendarEvent> userEvents = null;
    String calendarId = correspondenceService.getCorrespondingId(username, folderId);
    if (calendarId == null) {
      calendarId = CalendarConverterUtils.getCalendarId(folderId);
    }
    Calendar calendar = storage.getUserCalendar(username, calendarId);
    if (calendar != null) {
      List<String> calendarIds = new ArrayList<>();
      calendarIds.add(calendarId);
      userEvents = storage.getUserEventByCalendar(username, calendarIds);
    }
    return userEvents;
  }

  /**
   * Updates existing eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @throws Exception
   */
  public List<CalendarEvent> updateEvent(Appointment appointment, String username) throws Exception {
    return createOrUpdateEvent(appointment, username, false);
  }

  /**
   * Create non existing eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @throws Exception
   */
  public List<CalendarEvent> createEvent(Appointment appointment, String username) throws Exception {
    return createOrUpdateEvent(appointment, username, true);
  }

  /**
   * Creates or updates eXo Calendar Event.
   * 
   * @param appointment
   * @param folder
   * @param username
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> createOrUpdateEvent(Appointment appointment, String username) throws Exception {
    boolean isNew = correspondenceService.getCorrespondingId(username, appointment.getId().getUniqueId()) == null;
    if (!isNew) {
      CalendarEvent event = getEventByAppointmentId(username, appointment.getId().getUniqueId());
      if (event == null) {
        isNew = true;
        correspondenceService.deleteCorrespondingId(username, appointment.getId().getUniqueId());
      }
    }
    return createOrUpdateEvent(appointment, username, isNew);
  }

  /**
   * @param username
   * @param appointmentId
   * @return
   * @throws Exception
   */
  public CalendarEvent getEventByAppointmentId(String username, String appointmentId) throws Exception {
    String calEventId = correspondenceService.getCorrespondingId(username, appointmentId);
    CalendarEvent event = storage.getEvent(username, calEventId);
    if (event == null && calEventId != null) {
      correspondenceService.deleteCorrespondingId(username, appointmentId);
    }
    return event;
  }

  /**
   * @param eventNode
   * @return
   * @throws Exception
   */
  public CalendarEvent getExoEventByNode(Node eventNode) throws Exception {
    return storage.getEvent(eventNode);
  }

  /**
   * @param uuid
   * @return
   * @throws Exception
   */
  public String getExoEventMasterRecurenceByOriginalUUID(String uuid) throws Exception {
    Node node = storage.getSession(SessionProvider.createSystemProvider()).getNodeByUUID(uuid);
    if (node == null) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("No original recurrent node was found with UUID: " + uuid);
      }
      return null;
    } else {
      return node.getName();
    }
  }

  /**
   * @param username
   * @param calendar
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> getAllExoEvents(String username, Calendar calendar) throws Exception {
    List<String> calendarIds = Collections.singletonList(calendar.getId());
    return storage.getUserEventByCalendar(username, calendarIds);
  }

  /**
   * @param username
   * @param calendar
   * @param date
   * @return
   * @throws Exception
   */
  public List<CalendarEvent> findExoEventsModifiedSince(String username, Calendar calendar, Date date) throws Exception {
    Node calendarHome = storage.getUserCalendarHome(username);
    if (calendarHome.hasNode(calendar.getId())) {
      calendarHome = calendarHome.getNode(calendar.getId());
    }
    java.util.Calendar dateCalendar = java.util.Calendar.getInstance();
    dateCalendar.setTime(date);
    return getEventsByType(calendarHome, Calendar.TYPE_PRIVATE, dateCalendar);
  }

  private List<CalendarEvent> getEventsByType(Node calendarHome, int type, java.util.Calendar date) throws Exception {
    List<CalendarEvent> events = new ArrayList<>();
    QueryManager qm = calendarHome.getSession().getWorkspace().getQueryManager();
    Query query = qm.createQuery(
                                 "select * from exo:calendarEvent where (jcr:path like '" + calendarHome.getPath()
                                     + "/%') and (exo:lastModifiedDate > TIMESTAMP '" + ISO8601.format(date) + "')",
                                 Query.SQL);
    QueryResult result = query.execute();
    NodeIterator it = result.getNodes();
    CalendarEvent calEvent;
    while (it.hasNext()) {
      calEvent = storage.getEvent(it.nextNode());
      calEvent.setCalType(String.valueOf(type));
      events.add(calEvent);
    }
    return events;
  }

  public void updateModifiedDateOfEvent(String username, CalendarEvent event, Date lastModifiedTime) throws Exception {
    Node node = storage.getCalendarEventNode(username, event.getCalType(), event.getCalendarId(), event.getId());
    modifyUpdateDate(node, lastModifiedTime);
    if (event.getOriginalReference() != null && !event.getOriginalReference().isEmpty()) {
      Node masterNode = storage.getSession(SessionProvider.createSystemProvider()).getNodeByUUID(event.getOriginalReference());
      modifyUpdateDate(masterNode, lastModifiedTime);
    }
  }

  public CalendarEvent getEvent(String eventId, String username) throws Exception {
    return storage.getEvent(username, eventId);
  }

  private void modifyUpdateDate(Node node, Date lastModifiedTime) throws Exception {
    CalendarCreateUpdateAction.MODIFIED_DATE.set(lastModifiedTime.getTime());
    try {
      GregorianCalendar modifiedCalendar = new GregorianCalendar();
      modifiedCalendar.setTime(lastModifiedTime);
      if (!node.isNodeType(EXO_DATETIME_PROPERTY)) {
        if (node.canAddMixin(EXO_DATETIME_PROPERTY)) {
          node.addMixin(EXO_DATETIME_PROPERTY);
        }
        node.setProperty("exo:dateCreated", modifiedCalendar);
      }
      node.setProperty("exo:dateModified", modifiedCalendar);
      node.save();
    } finally {
      CalendarCreateUpdateAction.MODIFIED_DATE.set(null);
    }
  }

  private List<CalendarEvent> createOrUpdateEvent(Appointment appointment, String username, boolean isNew) throws Exception {
    Calendar calendar = getUserCalendar(username, appointment.getParentFolderId().getUniqueId());
    if (calendar == null) {
      LOG.warn("Attempting to synchronize an event without existing associated eXo Calendar.");
      return null;
    }
    List<CalendarEvent> updatedEvents = new ArrayList<>();

    if (appointment.getAppointmentType() != null) {
      switch (appointment.getAppointmentType()) {
      case Single: {
        CalendarEvent event = null;
        if (isNew) {
          event = new CalendarEvent();
          event.setId(null);
          event.setCalendarId(calendar.getId());
          updatedEvents.add(event);
        } else {
          event = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          if (event.getLastModified() == getLastModifiedDate(appointment).getTime()) {
            LOG.trace("IGNORE update eXo event '{}', modified dates are the same", appointment.getSubject());
            return updatedEvents;
          }
          updatedEvents.add(event);
          if (CalendarConverterUtils.verifyModifiedDatesConflict(event, appointment)) {
            if (LOG.isTraceEnabled()) {
              LOG.trace("Attempting to update eXo Event with Exchange Event, but modification date of eXo is after, ignore updating.");
            }
            return updatedEvents;
          }
        }

        if (LOG.isDebugEnabled()) {
          if (isNew) {
            LOG.debug("CREATE user calendar event: " + appointment.getSubject());
          } else {
            LOG.debug("UPDATE user calendar event: " + appointment.getSubject());
          }
        }

        CalendarConverterUtils.convertExchangeToExoEvent(event,
                                                         appointment,
                                                         username,
                                                         storage,
                                                         organizationService.getUserHandler());
        event.setRepeatType(CalendarEvent.RP_NOREPEAT);
        CalendarCreateUpdateAction.MODIFIED_DATE.set(getLastModifiedDate(appointment).getTime());
        try {
          storage.saveUserEvent(username, calendar.getId(), event, isNew);
        } finally {
          CalendarCreateUpdateAction.MODIFIED_DATE.set(null);
        }
        correspondenceService.setCorrespondingId(username, event.getId(), appointment.getId().getUniqueId());
      }
        break;
      case Exception:
        throw new IllegalStateException("The appointment is an exception occurence of this event >> '" + appointment.getSubject()
            + "'. start:" + appointment.getStart() + ", end : " + appointment.getEnd() + ", occurence: "
            + appointment.getAppointmentSequenceNumber());
      case RecurringMaster: {
        // Master recurring event
        CalendarEvent masterEvent = null;
        Date orginialStartDate = null;
        if (isNew) {
          masterEvent = new CalendarEvent();
          masterEvent.setId(null);
          updatedEvents.add(masterEvent);
        } else {
          masterEvent = getEventByAppointmentId(username, appointment.getId().getUniqueId());
          updatedEvents.add(masterEvent);
          orginialStartDate = masterEvent.getFromDateTime();
        }

        // there is a bug in Exchange modification time of the server:
        // Adding a recurrent Item + delete last occurence => last modified
        // date isn't updated. So we test here if the last occurence was deleted
        // or not
        //
        // Begin workaround
        boolean isLastOccurenceDeleted = false;
        appointment = Appointment.bind(appointment.getService(),
                                       appointment.getId(),
                                       new PropertySet(AppointmentSchema.Recurrence));
        if (appointment.getRecurrence().hasEnd()) {
          Date recEndDate = appointment.getRecurrence().getEndDate();

          appointment = Appointment.bind(appointment.getService(),
                                         appointment.getId(),
                                         new PropertySet(BasePropertySet.FirstClassProperties));
          if (recEndDate == null) {
            LOG.warn("Inconsistent data delivered by MS Exchange. The recurrent Event has end but end date is null: '"
                + appointment.getSubject() + "', start:" + appointment.getStart() + ", end : " + appointment.getEnd());
          } else {
            Appointment tmpAppointment = Appointment.bind(appointment.getService(),
                                                          appointment.getId(),
                                                          new PropertySet(AppointmentSchema.LastOccurrence));
            if (tmpAppointment.getLastOccurrence() == null) {
              LOG.warn("Can't find last occurence of recurrent Event : '" + appointment.getSubject() + "', start:"
                  + appointment.getStart() + ", end : " + appointment.getEnd());
            } else {
              isLastOccurenceDeleted = tmpAppointment.getLastOccurrence().getEnd().getTime() < recEndDate.getTime();

              if (isLastOccurenceDeleted && masterEvent.getExceptionIds() != null) {
                String pattern = EXCLUDE_ID_FORMAT_FIRST_CHARS.format(recEndDate);
                int i = 0;
                while (isLastOccurenceDeleted && i < masterEvent.getExceptionIds().size()) {
                  isLastOccurenceDeleted = !((String) masterEvent.getExceptionIds().toArray()[i]).startsWith(pattern);
                  i++;
                }
              }
            }
          }
        }
        appointment = Appointment.bind(appointment.getService(),
                                       appointment.getId(),
                                       new PropertySet(BasePropertySet.FirstClassProperties));
        // End workaround

        if (!isLastOccurenceDeleted && !isNew && CalendarConverterUtils.verifyModifiedDatesConflict(masterEvent, appointment)) {
          if (LOG.isTraceEnabled()) {
            LOG.trace("Attempting to update eXo Event with Exchange Event, but modification date of eXo is after, ignore updating.");
          }
          return updatedEvents;
        } else {
          if (LOG.isTraceEnabled()) {
            if (isNew) {
              LOG.debug("CREATE recurrent user calendar event: " + appointment.getSubject());
            } else {
              LOG.debug("UPDATE recurrent user calendar event: " + appointment.getSubject());
            }
          }

          masterEvent.setCalendarId(calendar.getId());
          CalendarConverterUtils.convertExchangeToExoMasterRecurringCalendarEvent(masterEvent,
                                                                                  appointment,
                                                                                  username,
                                                                                  storage,
                                                                                  organizationService.getUserHandler());
          if (isNew) {
            correspondenceService.setCorrespondingId(username, masterEvent.getId(), appointment.getId().getUniqueId());
          } else if (!CalendarConverterUtils.isSameDate(orginialStartDate, masterEvent.getFromDateTime())) {
            if (masterEvent.getExceptionIds() == null) {
              masterEvent.setExceptionIds(new ArrayList<String>());
            }
          }

          CalendarCreateUpdateAction.MODIFIED_DATE.set(getLastModifiedDate(appointment).getTime());
          try {
            storage.saveUserEvent(username, calendar.getId(), masterEvent, isNew);
          } finally {
            CalendarCreateUpdateAction.MODIFIED_DATE.set(null);
          }
        }
        List<CalendarEvent> exceptionalEventsToUpdate = new ArrayList<>();
        List<Appointment> occAppointments = new ArrayList<>();
        // Deleted execptional occurences events.
        List<CalendarEvent> toDeleteEvents =
                                           CalendarConverterUtils.convertExchangeToExoOccurenceEvent(masterEvent,
                                                                                                     exceptionalEventsToUpdate,
                                                                                                     occAppointments,
                                                                                                     appointment,
                                                                                                     username,
                                                                                                     storage,
                                                                                                     organizationService.getUserHandler(),
                                                                                                     correspondenceService);
        if (exceptionalEventsToUpdate != null && !exceptionalEventsToUpdate.isEmpty()) {
          CalendarCreateUpdateAction.MODIFIED_DATE.set(getLastModifiedDate(appointment).getTime());
          try {
            storage.updateOccurrenceEvent(calendar.getId(),
                                          calendar.getId(),
                                          masterEvent.getCalType(),
                                          masterEvent.getCalType(),
                                          exceptionalEventsToUpdate,
                                          username);
          } finally {
            CalendarCreateUpdateAction.MODIFIED_DATE.set(null);
          }

          // Set correspondance IDs
          Iterator<CalendarEvent> eventsIterator = exceptionalEventsToUpdate.iterator();
          Iterator<Appointment> occAppointmentIdIterator = occAppointments.iterator();
          while (eventsIterator.hasNext()) {
            CalendarEvent calendarEvent = eventsIterator.next();
            Appointment occAppointment = occAppointmentIdIterator.next();
            correspondenceService.setCorrespondingId(username, calendarEvent.getId(), occAppointment.getId().getUniqueId());
            updateModifiedDateOfEvent(username, calendarEvent, getLastModifiedDate(occAppointment));
          }
          updatedEvents.addAll(exceptionalEventsToUpdate);
        }
        if (toDeleteEvents != null && !toDeleteEvents.isEmpty()) {
          for (CalendarEvent calendarEvent : toDeleteEvents) {
            deleteEvent(username, calendarEvent);
          }
        }
      }
        break;
      case Occurrence:
        LOG.warn("The appointment is an occurence of this event >> '" + appointment.getSubject() + "'. start:"
            + appointment.getStart() + ", end : " + appointment.getEnd() + ", occurence: "
            + appointment.getAppointmentSequenceNumber());
      }
    }
    return updatedEvents;
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
}
