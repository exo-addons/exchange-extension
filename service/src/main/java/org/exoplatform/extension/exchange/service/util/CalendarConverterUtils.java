package org.exoplatform.extension.exchange.service.util;

import java.io.ByteArrayInputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Calendar;
import java.util.stream.Collectors;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.tika.mime.MimeTypes;

import org.exoplatform.calendar.service.*;
import org.exoplatform.calendar.service.impl.JCRDataStorage;
import org.exoplatform.commons.utils.ListAccess;
import org.exoplatform.extension.exchange.service.CorrespondenceService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.*;

import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.*;
import microsoft.exchange.webservices.data.core.enumeration.property.time.*;
import microsoft.exchange.webservices.data.core.enumeration.service.calendar.AppointmentType;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceObjectPropertyException;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.misc.TimeSpan;
import microsoft.exchange.webservices.data.property.complex.*;
import microsoft.exchange.webservices.data.property.complex.recurrence.DayOfTheWeekCollection;
import microsoft.exchange.webservices.data.property.complex.recurrence.pattern.Recurrence;
import microsoft.exchange.webservices.data.property.complex.recurrence.pattern.Recurrence.*;
import microsoft.exchange.webservices.data.util.DateTimeUtils;

/**
 * @author Boubaker Khanfir
 */
@SuppressWarnings("all")
public class CalendarConverterUtils {

  private static final Log                LOG                           = ExoLogger.getLogger(CalendarConverterUtils.class);

  public static final String              EXCHANGE_CALENDAR_NAME_PREFIX = "EXCH";

  public static final String              EXCHANGE_CALENDAR_ID_PREFIX   = "EXCH";

  public static final String              EXCHANGE_EVENT_ID_PREFIX      = "ExcangeEvent";

  public static final SimpleDateFormat    UTC_DATE_FORMAT               = new SimpleDateFormat("dd MMM yyyy HH:mm:ss");

  public static final SimpleDateFormat    RECURRENCE_ID_FORMAT          = new SimpleDateFormat("yyyyMMdd'T'HHmmss'Z'");

  private static final int                TIMEZONE_OFFSET_MILLIS        = TimeZone.getDefault().getRawOffset();

  // Reuse the object and save memory instead of instantiating this every call
  private static final ThreadLocal<Query> queryThreadLocal              = new ThreadLocal<>();

  private CalendarConverterUtils() {
  }

  /**
   * Converts from Exchange Calendar Event to eXo Calendar Event.
   * 
   * @param event
   * @param appointment
   * @param username
   * @param storage
   * @param userHandler
   * @throws Exception
   */
  public static void convertExchangeToExoEvent(CalendarEvent event,
                                               Appointment appointment,
                                               String username,
                                               JCRDataStorage storage,
                                               UserHandler userHandler) throws Exception {
    if (event.getId() == null || event.getId().isEmpty()) {
      event.setId(getEventId(appointment.getId().getUniqueId()));
    }
    if (event.getEventType() == null) {
      event.setEventType(CalendarEvent.TYPE_EVENT);
      event.setCalType("" + org.exoplatform.calendar.service.Calendar.TYPE_PRIVATE);
    }
    event.setLocation(appointment.getLocation());
    event.setSummary(appointment.getSubject());
    setEventStatus(event, appointment);
    setEventDates(event, appointment);
    setEventPriority(event, appointment);
    setEventCategory(event, appointment, username, storage);
    setEventParticipants(event, appointment, userHandler);
    setEventReminder(event, appointment, username);

    if (appointment.getSensitivity() != null && !appointment.getSensitivity().equals(Sensitivity.Normal)) {
      event.setPrivate(true);
    } else {
      event.setPrivate(false);
    }
    setEventAttachements(event, appointment);
    // This have to be last thing to load because of BAD EWS API impl
    setEventDescription(event, appointment);
  }

  /**
   * Converts from Exchange Calendar Recurring Master Event to eXo Calendar
   * Event.
   * 
   * @param event
   * @param appointment
   * @param username
   * @param storage
   * @param userHandler
   * @throws Exception
   */
  @SuppressWarnings("deprecation")
  public static void convertExchangeToExoMasterRecurringCalendarEvent(CalendarEvent event,
                                                                      Appointment appointment,
                                                                      String username,
                                                                      JCRDataStorage storage,
                                                                      UserHandler userHandler) throws Exception {
    convertExchangeToExoEvent(event, appointment, username, storage, userHandler);
    appointment = Appointment.bind(appointment.getService(), appointment.getId(), new PropertySet(AppointmentSchema.Recurrence));
    Recurrence recurrence = appointment.getRecurrence();
    if (recurrence instanceof DailyPattern) {
      event.setRepeatType(CalendarEvent.RP_DAILY);
    } else if (recurrence instanceof WeeklyPattern) {
      event.setRepeatType(CalendarEvent.RP_WEEKLY);
      DayOfTheWeekCollection dayOfTheWeekCollection = ((WeeklyPattern) recurrence).getDaysOfTheWeek();
      if (dayOfTheWeekCollection != null && dayOfTheWeekCollection.getCount() > 0) {
        String[] days = new String[7];
        Iterator<DayOfTheWeek> iterator = dayOfTheWeekCollection.iterator();
        ITERATE_WHILE: while (iterator.hasNext()) {
          DayOfTheWeek dayOfTheWeek = (DayOfTheWeek) iterator.next();
          switch (dayOfTheWeek) {
          case Monday:
            days[0] = CalendarEvent.RP_WEEKLY_BYDAY[0];
            break;
          case Tuesday:
            days[1] = CalendarEvent.RP_WEEKLY_BYDAY[1];
            break;
          case Wednesday:
            days[2] = CalendarEvent.RP_WEEKLY_BYDAY[2];
            break;
          case Thursday:
            days[3] = CalendarEvent.RP_WEEKLY_BYDAY[3];
            break;
          case Friday:
            days[4] = CalendarEvent.RP_WEEKLY_BYDAY[4];
            break;
          case Saturday:
            days[5] = CalendarEvent.RP_WEEKLY_BYDAY[5];
            break;
          case Sunday:
            days[6] = CalendarEvent.RP_WEEKLY_BYDAY[6];
            break;
          case Weekday:
            days = Arrays.copyOfRange(CalendarEvent.RP_WEEKLY_BYDAY, 0, 4);
            event.setRepeatType(CalendarEvent.RP_WORKINGDAYS);
            break ITERATE_WHILE;
          case WeekendDay:
            days = Arrays.copyOfRange(CalendarEvent.RP_WEEKLY_BYDAY, 5, 6);
            event.setRepeatType(CalendarEvent.RP_WEEKEND);
            break ITERATE_WHILE;
          case Day:
            days = Arrays.copyOfRange(CalendarEvent.RP_WEEKLY_BYDAY, 0, 6);
            break ITERATE_WHILE;
          }
        }
        List<String> daysList = new ArrayList<String>();
        for (String day : days) {
          if (day != null) {
            daysList.add(day);
          }
        }
        event.setRepeatByDay(daysList.toArray(new String[0]));
      }
    } else if (recurrence instanceof RelativeMonthlyPattern) {
      event.setRepeatType(CalendarEvent.RP_MONTHLY);

      DayOfTheWeekIndex dayOfTheWeekIndex = ((RelativeMonthlyPattern) recurrence).getDayOfTheWeekIndex();
      int exoIndex = (dayOfTheWeekIndex.ordinal() + 2) % 6 - 1;
      String dayPrefix = null;

      switch (((RelativeMonthlyPattern) recurrence).getDayOfTheWeek()) {
      case Monday:
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[0];
        break;
      case Tuesday:
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[1];
        break;
      case Wednesday:
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[2];
        break;
      case Thursday:
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[3];
        break;
      case Friday:
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[4];
        break;
      case Saturday:
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[5];
        break;
      case Sunday:
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[6];
        break;
      default:
        LOG.error("Uknown day of the week '" + ((RelativeMonthlyPattern) recurrence).getDayOfTheWeek() + "', for recurrence '"
            + appointment.getSubject() + "', Monday will be set.");
        dayPrefix = CalendarEvent.RP_WEEKLY_BYDAY[0];
        break;
      }

      event.setRepeatByDay(new String[] { exoIndex + dayPrefix });
      event.setRepeatByMonthDay(null);
    } else if (recurrence instanceof MonthlyPattern) {
      event.setRepeatType(CalendarEvent.RP_MONTHLY);
      event.setRepeatByDay(null);
      int dayOfMonth = 0;
      try {
        dayOfMonth = ((MonthlyPattern) recurrence).getDayOfMonth();
      } catch (Exception e) {
        dayOfMonth = recurrence.getStartDate().getDate();
      }
      event.setRepeatByMonthDay(new long[] { dayOfMonth });
    } else if (recurrence instanceof YearlyPattern) {
      event.setRepeatType(CalendarEvent.RP_YEARLY);

      event.setRepeatInterval(1);
      Calendar tempCalendar = Calendar.getInstance();
      tempCalendar.set(Calendar.MONTH, ((YearlyPattern) recurrence).getMonth().ordinal());
      int dayOfMonth = 0;
      try {
        dayOfMonth = ((MonthlyPattern) recurrence).getDayOfMonth();
      } catch (Exception e) {
        dayOfMonth = recurrence.getStartDate().getDate();
      }
      tempCalendar.set(Calendar.DAY_OF_MONTH, dayOfMonth);
      int dayOfYear = tempCalendar.get(Calendar.DAY_OF_YEAR);

      event.setRepeatByYearDay(new long[] { dayOfYear });
    } else if (recurrence instanceof RelativeYearlyPattern) {
      LOG.error("Cannot manage events of type RelativeYearlyPattern.");
      event.setRepeatType(CalendarEvent.RP_YEARLY);
    }
    if (recurrence instanceof IntervalPattern) {
      if (((IntervalPattern) recurrence).getInterval() > 0) {
        event.setRepeatInterval(((IntervalPattern) recurrence).getInterval());
      }
    }
    Date endDate = recurrence.getEndDate();
    if (endDate != null) {
      Calendar calendar = getCalendarInstance(endDate);
      setEndOfDay(calendar);
      event.setRepeatUntilDate(calendar.getTime());
    }
    if (recurrence.getNumberOfOccurrences() != null) {
      event.setRepeatCount(recurrence.getNumberOfOccurrences());
    }
  }

  /**
   * Converts from Exchange Calendar Exceptional Occurence Event to eXo Calendar
   * Event and return the list of deleted and updated elements.
   * 
   * @param masterEvent
   * @param updatedEvents empty list that will be updated by modified occurences
   * @param occurenceAppointments
   * @param masterAppointment
   * @param username
   * @param storage
   * @param userHandler
   * @return
   * @throws Exception
   */
  public static List<CalendarEvent> convertExchangeToExoOccurenceEvent(CalendarEvent masterEvent,
                                                                       List<CalendarEvent> updatedEvents,
                                                                       List<Appointment> occurenceAppointments,
                                                                       Appointment masterAppointment,
                                                                       String username,
                                                                       JCRDataStorage storage,
                                                                       UserHandler userHandler,
                                                                       CorrespondenceService correspondenceService) throws Exception {
    masterAppointment = Appointment.bind(masterAppointment.getService(),
                                         masterAppointment.getId(),
                                         new PropertySet(AppointmentSchema.ModifiedOccurrences));
    {
      OccurrenceInfoCollection occurrenceInfoCollection = masterAppointment.getModifiedOccurrences();
      if (occurrenceInfoCollection != null && occurrenceInfoCollection.getCount() > 0) {
        for (OccurrenceInfo occurrenceInfo : occurrenceInfoCollection) {
          Appointment occurenceAppointment = Appointment.bind(masterAppointment.getService(),
                                                              occurrenceInfo.getItemId(),
                                                              new PropertySet(BasePropertySet.FirstClassProperties));

          String exoId = correspondenceService.getCorrespondingId(username, occurenceAppointment.getId().getUniqueId());
          CalendarEvent tmpEvent = null;
          if (!StringUtils.isEmpty(exoId)) {
            tmpEvent = storage.getEvent(username, exoId);
          }
          if (tmpEvent == null) {
            tmpEvent = getOccurenceOfDate(username, storage, masterEvent, occurrenceInfo.getOriginalStart());
          }

          if (tmpEvent != null && verifyModifiedDatesConflict(tmpEvent, occurenceAppointment)) {
            if (LOG.isDebugEnabled()) {
              LOG.debug("CONFLICT: Attempting to update eXo Event with Exchange Event, but modification date of eXo is after, ignore updating.");
            }
            continue;
          }
          if (tmpEvent == null || tmpEvent.getIsExceptionOccurrence() == null || !tmpEvent.getIsExceptionOccurrence()) {
            tmpEvent = new CalendarEvent();
            convertExchangeToExoEvent(tmpEvent, occurenceAppointment, username, storage, userHandler);
            tmpEvent.setRecurrenceId(RECURRENCE_ID_FORMAT.format(tmpEvent.getFromDateTime()));
            tmpEvent.setRepeatType(CalendarEvent.RP_NOREPEAT);
            tmpEvent.setId(masterEvent.getId());
            tmpEvent.setCalendarId(masterEvent.getCalendarId());
            if (LOG.isDebugEnabled()) {
              LOG.debug("CREATE exo calendar Occurence event: " + tmpEvent.getSummary() + ", with recurence id: "
                  + tmpEvent.getRecurrenceId());
            }
          } else {
            if (LOG.isDebugEnabled()) {
              LOG.debug("UPDATE exo calendar Occurence event: " + tmpEvent.getSummary() + ", with recurence id: "
                  + tmpEvent.getRecurrenceId());
            }
            convertExchangeToExoEvent(tmpEvent, occurenceAppointment, username, storage, userHandler);
          }
          updatedEvents.add(tmpEvent);
          occurenceAppointments.add(occurenceAppointment);
        }
      }
    }
    masterAppointment = Appointment.bind(masterAppointment.getService(),
                                         masterAppointment.getId(),
                                         new PropertySet(AppointmentSchema.DeletedOccurrences));

    List<CalendarEvent> calendarEvents = new ArrayList<CalendarEvent>();
    DeletedOccurrenceInfoCollection deletedOccurrenceInfoCollection = masterAppointment.getDeletedOccurrences();
    if (deletedOccurrenceInfoCollection != null && deletedOccurrenceInfoCollection.getCount() > 0) {
      for (DeletedOccurrenceInfo occurrenceInfo : deletedOccurrenceInfoCollection) {
        CalendarEvent toDeleteEvent = getOccurenceOfDate(username, storage, masterEvent, occurrenceInfo.getOriginalStart());
        if (toDeleteEvent == null) {
          continue;
        }

        String appId = correspondenceService.getCorrespondingId(username, toDeleteEvent.getId());
        Appointment appointment = null;
        try {
          appointment = Appointment.bind(masterAppointment.getService(), ItemId.getItemIdFromString(appId));
        } catch (Exception e) {
        }
        if (appointment == null || toDeleteEvent.getIsExceptionOccurrence() == null
            || !toDeleteEvent.getIsExceptionOccurrence()) {
          calendarEvents.add(toDeleteEvent);
        }
      }
    }
    return calendarEvents;
  }

  /**
   * @param event eXo Calendar event
   * @param item Exchange item
   * @return
   * @throws Exception
   */
  @SuppressWarnings("deprecation")
  public static boolean verifyModifiedDatesConflict(CalendarEvent event, Appointment appointment) throws Exception {
    if (event.getLastUpdatedTime() == null) {
      return false;
    } else if (appointment.getLastModifiedTime() == null) {
      return true;
    }
    return event.getLastModified() > appointment.getLastModifiedTime().getTime();
  }

  /**
   * Converts from Exchange Calendar Event to eXo Calendar Event.
   * 
   * @param calendarEvent
   * @param appointment
   * @param username
   * @param calendarService
   * @throws Exception
   */
  public static void convertExoToExchangeEvent(Appointment appointment,
                                               CalendarEvent calendarEvent,
                                               String username,
                                               UserHandler userHandler) throws Exception {
    setAppointmentStatus(appointment, calendarEvent);
    setAppointmentDates(appointment, calendarEvent);
    setAppointmentPriority(appointment, calendarEvent);
    setAppointmentCategory(appointment, calendarEvent);
    setAppointmentAttendees(appointment, calendarEvent, userHandler, username);
    setAppointmentReminder(appointment, calendarEvent);

    appointment.setLocation(calendarEvent.getLocation());
    appointment.setSubject(calendarEvent.getSummary());
    if (calendarEvent.isPrivate()) {
      appointment.setSensitivity(Sensitivity.Private);
    } else {
      appointment.setSensitivity(Sensitivity.Normal);
    }
    setAppointmentAttachements(appointment, calendarEvent);

    // This have to be last thing to load because of BAD EWS API impl
    setApoinementSummary(appointment, calendarEvent);
  }

  /**
   * Converts from Exchange Calendar Recurring Master Event to eXo Calendar
   * Event.
   * 
   * @param event
   * @param appointment
   * @param username
   * @param calendarService
   * @return list of occurences to delete
   * @throws Exception
   */
  public static List<Appointment> convertExoToExchangeMasterRecurringCalendarEvent(Appointment appointment,
                                                                                   CalendarEvent event,
                                                                                   String username,
                                                                                   UserHandler userHandler) throws Exception {
    List<Appointment> toDeleteOccurences = null;

    convertExoToExchangeEvent(appointment, event, username, userHandler);

    String repeatType = event.getRepeatType();
    assert repeatType != null && !repeatType.equals(CalendarEvent.RP_NOREPEAT);
    Recurrence recurrence = null;
    if (repeatType.equals(CalendarEvent.RP_DAILY)) {
      recurrence = new Recurrence.DailyPattern();
    } else {
      long repeatInterval = event.getRepeatInterval();
      if (repeatInterval == 0) {
        repeatInterval = 1;
      }
      if (repeatType.equals(CalendarEvent.RP_WEEKLY)) {
        List<DayOfTheWeek> daysOfTheWeek = new ArrayList<DayOfTheWeek>();
        String[] repeatDays = event.getRepeatByDay();
        if (repeatDays != null) {
          for (String repeatDay : repeatDays) {
            if (StringUtils.isEmpty(repeatDay)) {
              continue;
            } else if (repeatDay.equals(CalendarEvent.RP_WEEKLY_BYDAY[0])) {
              daysOfTheWeek.add(DayOfTheWeek.Monday);
            } else if (repeatDay.equals(CalendarEvent.RP_WEEKLY_BYDAY[1])) {
              daysOfTheWeek.add(DayOfTheWeek.Tuesday);
            } else if (repeatDay.equals(CalendarEvent.RP_WEEKLY_BYDAY[2])) {
              daysOfTheWeek.add(DayOfTheWeek.Wednesday);
            } else if (repeatDay.equals(CalendarEvent.RP_WEEKLY_BYDAY[3])) {
              daysOfTheWeek.add(DayOfTheWeek.Thursday);
            } else if (repeatDay.equals(CalendarEvent.RP_WEEKLY_BYDAY[4])) {
              daysOfTheWeek.add(DayOfTheWeek.Friday);
            } else if (repeatDay.equals(CalendarEvent.RP_WEEKLY_BYDAY[5])) {
              daysOfTheWeek.add(DayOfTheWeek.Saturday);
            } else if (repeatDay.equals(CalendarEvent.RP_WEEKLY_BYDAY[6])) {
              daysOfTheWeek.add(DayOfTheWeek.Sunday);
            }
          }
        }
        recurrence = new Recurrence.WeeklyPattern(event.getFromDateTime(),
                                                  (int) repeatInterval,
                                                  daysOfTheWeek.toArray(new DayOfTheWeek[0]));
      } else if (repeatType.equals(CalendarEvent.RP_MONTHLY)) {
        long[] repeatByMonthDay = event.getRepeatByMonthDay();
        if ((repeatByMonthDay == null || repeatByMonthDay.length == 0)
            && (event.getRepeatByDay() == null || event.getRepeatByDay().length == 0)) {
          repeatByMonthDay = new long[] { event.getFromDateTime().getDate() };
        }
        if (repeatByMonthDay != null && repeatByMonthDay.length > 0) {
          recurrence = new Recurrence.MonthlyPattern(event.getFromDateTime(), (int) repeatInterval, (int) repeatByMonthDay[0]);
        } else if ((event.getRepeatByDay() != null || event.getRepeatByDay().length > 0)) {
          String repeatByDay = event.getRepeatByDay()[0];
          int weekIndex = Integer.parseInt(repeatByDay.substring(0, 1));
          String dayPrefix = repeatByDay.substring(1);

          DayOfTheWeek dayOfTheWeek = null;
          DayOfTheWeekIndex dayOfTheWeekIndex = null;

          if (dayPrefix.equals(CalendarEvent.RP_WEEKLY_BYDAY[0])) {
            dayOfTheWeek = DayOfTheWeek.Monday;
          } else if (dayPrefix.equals(CalendarEvent.RP_WEEKLY_BYDAY[1])) {
            dayOfTheWeek = DayOfTheWeek.Tuesday;
          } else if (dayPrefix.equals(CalendarEvent.RP_WEEKLY_BYDAY[2])) {
            dayOfTheWeek = DayOfTheWeek.Wednesday;
          } else if (dayPrefix.equals(CalendarEvent.RP_WEEKLY_BYDAY[3])) {
            dayOfTheWeek = DayOfTheWeek.Thursday;
          } else if (dayPrefix.equals(CalendarEvent.RP_WEEKLY_BYDAY[4])) {
            dayOfTheWeek = DayOfTheWeek.Friday;
          } else if (dayPrefix.equals(CalendarEvent.RP_WEEKLY_BYDAY[5])) {
            dayOfTheWeek = DayOfTheWeek.Saturday;
          } else if (dayPrefix.equals(CalendarEvent.RP_WEEKLY_BYDAY[6])) {
            dayOfTheWeek = DayOfTheWeek.Sunday;
          } else {
            LOG.error("Can't get day of the week name from this prefix: '" + dayPrefix + "'. Monday will be used");
            dayOfTheWeek = DayOfTheWeek.Monday;
          }

          switch (weekIndex) {
          case -1:
            dayOfTheWeekIndex = DayOfTheWeekIndex.Last;
            break;
          case 1:
            dayOfTheWeekIndex = DayOfTheWeekIndex.First;
            break;
          case 2:
            dayOfTheWeekIndex = DayOfTheWeekIndex.Second;
            break;
          case 3:
            dayOfTheWeekIndex = DayOfTheWeekIndex.Third;
            break;
          case 4:
            dayOfTheWeekIndex = DayOfTheWeekIndex.Fourth;
            break;
          default:
            LOG.error("Can't get week index from this number: '" + weekIndex + "'. First week will be used as default value");
            dayOfTheWeekIndex = DayOfTheWeekIndex.First;
            break;
          }
          recurrence = new Recurrence.RelativeMonthlyPattern(event.getFromDateTime(),
                                                             (int) repeatInterval,
                                                             dayOfTheWeek,
                                                             dayOfTheWeekIndex);
        }
      } else if (repeatType.equals(CalendarEvent.RP_YEARLY)) {
        Calendar tempCalendar = Calendar.getInstance();
        if (event.getRepeatByYearDay() != null && event.getRepeatByYearDay().length > 0) {
          tempCalendar.set(Calendar.DAY_OF_YEAR, (int) event.getRepeatByYearDay()[0]);
        } else {
          tempCalendar.setTime(event.getFromDateTime());
        }

        int monthNumber = tempCalendar.get(Calendar.MONTH);
        int dayOfMonth = tempCalendar.get(Calendar.DAY_OF_MONTH);

        Month month = Month.values()[monthNumber];
        recurrence = new Recurrence.YearlyPattern(event.getFromDateTime(), month, dayOfMonth);
      } else if (repeatType.equals(CalendarEvent.RP_WORKINGDAYS)) {
        recurrence = new Recurrence.WeeklyPattern(event.getFromDateTime(), (int) repeatInterval, DayOfTheWeek.Weekday);
      } else if (repeatType.equals(CalendarEvent.RP_WEEKEND)) {
        recurrence = new Recurrence.WeeklyPattern(event.getFromDateTime(), (int) repeatInterval, DayOfTheWeek.WeekendDay);
      }
    }

    recurrence.setStartDate(event.getFromDateTime());

    if (event.getRepeatUntilDate() == null && event.getRepeatCount() < 1) {
      recurrence.neverEnds();
    } else if (event.getRepeatUntilDate() != null) {
      recurrence.setEndDate(event.getRepeatUntilDate());
    } else {
      recurrence.setNumberOfOccurrences((int) event.getRepeatCount());
    }

    appointment.setRecurrence(recurrence);

    if (event.getExceptionIds() != null && event.getExceptionIds().size() > 0) {
      toDeleteOccurences = calculateOccurences(username, appointment, event);

      int nbOccurences = recurrence.getNumberOfOccurrences() == null ? 0 : recurrence.getNumberOfOccurrences();
      int deletedAppointmentOccurences = 0;
      try {
        deletedAppointmentOccurences = appointment.getDeletedOccurrences().getCount();
      } catch (Exception e) {
        try {
          appointment = Appointment.bind(appointment.getService(),
                                         appointment.getId(),
                                         new PropertySet(BasePropertySet.FirstClassProperties));
          deletedAppointmentOccurences = appointment.getDeletedOccurrences().getCount();
        } catch (Exception e2) {
          deletedAppointmentOccurences = 0;
        }
      }
      if (nbOccurences != 0 && (nbOccurences - deletedAppointmentOccurences - toDeleteOccurences.size()) == 0) {
        toDeleteOccurences.clear();
        toDeleteOccurences.add(appointment);
      }
    }
    return toDeleteOccurences;
  }

  /**
   * Converts from Exchange Calendar Exceptional Occurence Event to eXo Calendar
   * Event.
   * 
   * @param masterEvent
   * @param listEvent
   * @param masterAppointment
   * @param username
   * @param calendarService
   * @return
   * @throws Exception
   */
  public static void convertExoToExchangeOccurenceEvent(Appointment occAppointment,
                                                        CalendarEvent occEvent,
                                                        String username,
                                                        UserHandler userHandler) throws Exception {
    convertExoToExchangeEvent(occAppointment, occEvent, username, userHandler);
  }

  /**
   * Converts Exchange Calendar Name to eXo Calendar Name by adding a prefix.
   * 
   * @param calendarName
   * @return
   */
  public static String getCalendarName(String calendarName) {
    return EXCHANGE_CALENDAR_NAME_PREFIX + "-" + calendarName;
  }

  /**
   * Converts Exchange Calendar Name to eXo Calendar Id by adding a prefix and
   * hash coding the original Id.
   * 
   * @param folderId
   * @return
   */
  public static String getCalendarId(String folderId) {
    return EXCHANGE_CALENDAR_ID_PREFIX + "-" + folderId.hashCode();
  }

  /**
   * Checks if Passed eXo Calendar Id becomes from the synchronization with
   * exchange, by testing if the prefix exists or not.
   * 
   * @param calendarId
   * @return
   */
  public static boolean isExchangeCalendarId(String calendarId) {
    return calendarId != null && calendarId.startsWith(EXCHANGE_CALENDAR_ID_PREFIX);
  }

  /**
   * Checks if Passed eXo Calendar Id becomes from the synchronization with
   * exchange, by testing if the prefix exists or not.
   * 
   * @param calendarId
   * @return
   */
  public static boolean isExchangeEventId(String eventId) {
    return eventId != null && eventId.startsWith(EXCHANGE_EVENT_ID_PREFIX);
  }

  /**
   * Converts Exchange Calendar Event Id to eXo Calendar Event Id
   * 
   * @param appointmentId
   * @return
   * @throws Exception
   */
  public static String getEventId(String appointmentId) throws Exception {
    return EXCHANGE_EVENT_ID_PREFIX + appointmentId.hashCode();
  }

  /**
   * Compares two dates.
   * 
   * @param value1
   * @param value2
   * @return true if same
   */
  public static boolean isSameDate(Date value1, Date value2) {
    Calendar date1 = Calendar.getInstance();
    date1.setTime(value1);
    Calendar date2 = Calendar.getInstance();
    date2.setTime(value2);
    return isSameDate(date1, date2);
  }

  public static boolean isAllDayEvent(CalendarEvent eventCalendar) {
    Calendar cal1 = Calendar.getInstance();
    cal1.setLenient(false);
    Calendar cal2 = Calendar.getInstance();
    cal2.setLenient(false);

    cal1.setTime(eventCalendar.getFromDateTime());
    cal2.setTime(eventCalendar.getToDateTime());
    return (cal1.get(Calendar.HOUR_OF_DAY) == 0 && cal1.get(Calendar.MINUTE) == 0
        && cal2.get(Calendar.HOUR_OF_DAY) == cal2.getActualMaximum(Calendar.HOUR_OF_DAY)
        && cal2.get(Calendar.MINUTE) == cal2.getActualMaximum(Calendar.MINUTE));
  }

  public static Appointment getAppointmentOccurence(Appointment masterAppointment, String recurrenceId) throws Exception {
    Appointment appointment = null;
    Date occDate = CalendarConverterUtils.RECURRENCE_ID_FORMAT.parse(recurrenceId);
    {
      Calendar calendar = Calendar.getInstance();
      calendar.setTime(occDate);
      calendar.set(Calendar.HOUR_OF_DAY, 0);
      calendar.set(Calendar.MINUTE, 0);
      calendar.set(Calendar.SECOND, 0);
      calendar.set(Calendar.MILLISECOND, 0);
      occDate = calendar.getTime();
    }
    int i = 1;
    Date endDate = masterAppointment.getRecurrence().getEndDate();
    if (endDate != null && occDate.getTime() > endDate.getTime()) {
      return null;
    }
    Integer nbOccurences = masterAppointment.getRecurrence().getNumberOfOccurrences();

    Calendar indexCalendar = Calendar.getInstance();
    indexCalendar.setTime(masterAppointment.getRecurrence().getStartDate());

    boolean continueSearch = true;
    while (continueSearch && (nbOccurences == null || i <= nbOccurences)) {
      Appointment tmpAppointment = null;
      try {
        tmpAppointment = Appointment.bindToOccurrence(masterAppointment.getService(),
                                                      masterAppointment.getId(),
                                                      i,
                                                      new PropertySet(AppointmentSchema.Start));
        Date date = tmpAppointment.getStart();
        if (CalendarConverterUtils.isSameDate(occDate, date)) {
          appointment = Appointment.bindToOccurrence(masterAppointment.getService(),
                                                     masterAppointment.getId(),
                                                     i,
                                                     new PropertySet(BasePropertySet.FirstClassProperties));
          continueSearch = false;
        }
        indexCalendar.setTime(date);
      } catch (Exception e) {
        // increment date
        indexCalendar.add(Calendar.DATE, 1);
      }
      i++;
      if (continueSearch
          && (occDate.before(indexCalendar.getTime()) || (endDate != null && indexCalendar.getTime().after(endDate)))) {
        continueSearch = false;
      }
    }
    return appointment;
  }

  public static CalendarEvent getOccurenceOfDate(String username,
                                                 JCRDataStorage storage,
                                                 CalendarEvent masterEvent,
                                                 Date originalStart) throws Exception {
    Date date = originalStart;
    String recurenceId = RECURRENCE_ID_FORMAT.format(date);
    List<CalendarEvent> exceptionEvents = storage.getExceptionEvents(username, masterEvent);
    if (exceptionEvents != null && !exceptionEvents.isEmpty()) {
      for (CalendarEvent calendarEvent : exceptionEvents) {
        if (calendarEvent.getRecurrenceId().equals(recurenceId)) {
          return calendarEvent;
        }
      }
    }

    Calendar from = Calendar.getInstance();
    from.setTime(date);
    from.set(Calendar.HOUR_OF_DAY, 0);
    from.set(Calendar.MINUTE, 0);
    from.set(Calendar.SECOND, 0);
    from.set(Calendar.MILLISECOND, 0);

    Calendar to = Calendar.getInstance();
    to.setTime(date);
    setEndOfDay(to);

    CalendarSetting calendarSetting = storage.getCalendarSetting(username);
    Map<String, CalendarEvent> map = storage.getOccurrenceEvents(masterEvent, from, to, calendarSetting.getTimeZone());
    CalendarEvent occEvent = null;
    if (map != null && !map.isEmpty()) {
      if (map.size() == 1) {
        occEvent = map.values().iterator().next();
      } else {
        LOG.error("Error while deleting from eXo an occurence already deleted from Exchange '" + masterEvent.getSummary()
            + "' in date: '" + date + "'");
      }
    }
    return occEvent;
  }

  private static void setEndOfDay(Calendar to) {
    to.set(Calendar.HOUR_OF_DAY, to.getActualMaximum(Calendar.HOUR_OF_DAY));
    to.set(Calendar.MINUTE, to.getActualMaximum(Calendar.MINUTE));
    to.set(Calendar.SECOND, to.getActualMaximum(Calendar.SECOND));
    to.set(Calendar.MILLISECOND, to.getActualMaximum(Calendar.MILLISECOND));
  }

  private static void setAppointmentReminder(Appointment appointment, CalendarEvent calendarEvent) throws Exception {
    appointment.setIsReminderSet(false);
    List<Reminder> reminders = calendarEvent.getReminders();
    if (reminders != null) {
      for (Reminder reminder : reminders) {
        appointment.setIsReminderSet(true);
        appointment.setReminderMinutesBeforeStart((int) reminder.getAlarmBefore());
        appointment.setReminderDueBy(reminder.getFromDateTime());
      }
    }
  }

  private static List<Appointment> calculateOccurences(String username,
                                                       Appointment masterAppointment,
                                                       CalendarEvent event) throws Exception {
    List<Appointment> toDeleteOccurence = new ArrayList<Appointment>();
    Collection<String> excludedRecurenceIds = event.getExceptionIds();
    if (excludedRecurenceIds != null) {
      for (String excludedRecurenceId : excludedRecurenceIds) {
        if (excludedRecurenceId.isEmpty()) {
          continue;
        }
        Appointment occAppointment = getAppointmentOccurence(masterAppointment, excludedRecurenceId);

        if (occAppointment != null) {
          toDeleteOccurence.add(occAppointment);
        }
      }
    }
    return toDeleteOccurence;
  }

  /**
   * Converts Exchange Calendar Category Name to eXo Calendar Name
   * 
   * @param categoryName
   * @return
   */
  private static String getCategoryName(String categoryName) {
    return /* EXCHANGE_CALENDAR_NAME_PREFIX + "-" + */categoryName;
  }

  private static boolean isSameDate(java.util.Calendar date1, java.util.Calendar date2) {
    return (date1.get(java.util.Calendar.DATE) == date2.get(java.util.Calendar.DATE)
        && date1.get(java.util.Calendar.MONTH) == date2.get(java.util.Calendar.MONTH)
        && date1.get(java.util.Calendar.YEAR) == date2.get(java.util.Calendar.YEAR));
  }

  private static void setAppointmentAttendees(Appointment appointment,
                                              CalendarEvent calendarEvent,
                                              UserHandler userHandler,
                                              String username) throws ServiceLocalException {
    AttendeeCollection attendees = appointment.getRequiredAttendees();
    attendees.clear();
    computeAttendies(userHandler, username, attendees, calendarEvent.getParticipant());

    attendees = appointment.getOptionalAttendees();
    attendees.clear();
    computeAttendies(userHandler, username, attendees, calendarEvent.getInvitation());
  }

  private static void computeAttendies(UserHandler userHandler,
                                       String username,
                                       AttendeeCollection attendees,
                                       String[] participants) {
    if (participants != null && participants.length > 0) {
      for (String partacipant : participants) {
        if (partacipant == null || partacipant.isEmpty() || partacipant.equals(username)) {
          continue;
        }
        try {
          User user = userHandler.findUserByName(partacipant);
          if (!containsAttendee(attendees, user.getEmail())) {
            Attendee attendee = new Attendee(user.getDisplayName(), user.getEmail());
            attendees.add(attendee);
          }
        } catch (Exception e) {
          Attendee attendee = null;
          if (partacipant.contains("@")) {
            attendee = new Attendee(partacipant.split("@")[0], partacipant);
          } else {
            attendee = new Attendee(partacipant, null);
          }
          attendees.add(attendee);
          if (LOG.isTraceEnabled()) {
            LOG.warn("Partacipant '" + partacipant + "' wasn't found in eXo Organization.");
          }
        }
      }
    }
  }

  private static boolean containsAttendee(AttendeeCollection attendees, String email) {
    for (Attendee attendee : attendees) {
      if (attendee.getAddress().equals(email)) {
        return true;
      }
    }
    return false;
  }

  private static void setEventParticipants(CalendarEvent calendarEvent,
                                           Appointment appointment,
                                           UserHandler userHandler) throws ServiceLocalException {
    Query query = queryThreadLocal.get();
    if (query == null) {
      query = new Query();
      queryThreadLocal.set(query);
    }
    List<String> participants = new ArrayList<String>();
    addEventPartacipants(appointment.getRequiredAttendees(), userHandler, query, participants);
    addEventPartacipants(appointment.getOptionalAttendees(), userHandler, query, participants);
    addEventPartacipants(appointment.getResources(), userHandler, query, participants);
    if (participants.size() > 0) {
      calendarEvent.setParticipant(participants.toArray(new String[0]));
      List<String> participantsStatuses =
                                        participants.stream().map(participant -> participant + ":").collect(Collectors.toList());

      calendarEvent.setParticipantStatus(participantsStatuses.toArray(new String[0]));
    }
  }

  private static void addEventPartacipants(AttendeeCollection attendeeCollection,
                                           UserHandler userHandler,
                                           Query query,
                                           List<String> participants) throws ServiceLocalException {
    if (attendeeCollection != null && attendeeCollection.getItems() != null && !attendeeCollection.getItems().isEmpty()) {
      for (Attendee attendee : attendeeCollection.getItems()) {
        if (attendee.getAddress() != null && !attendee.getAddress().isEmpty()) {
          String username = getPartacipantUserName(userHandler, query, attendee);
          if (username == null) {
            if (LOG.isTraceEnabled()) {
              LOG.trace("Event partacipant was not found, email = " + attendee.getAddress());
            }
            username = attendee.getAddress() == null ? attendee.getName() : attendee.getAddress();
            if (username == null) {
              LOG.warn("No user found for attendee: " + attendee);
              continue;
            }
          }
          participants.add(username);
        }
      }
    }
  }

  private static String getPartacipantUserName(UserHandler userHandler, Query query, Attendee attendee) {
    String username = null;
    query.setEmail(attendee.getAddress());
    try {
      ListAccess<User> listAccess = userHandler.findUsersByQuery(query);
      if (listAccess == null || listAccess.getSize() == 0) {
        if (LOG.isTraceEnabled()) {
          LOG.trace("User with email '" + attendee.getAddress() + "' was not found in eXo.");
        }
      } else if (listAccess.getSize() > 1) {
        if (LOG.isTraceEnabled()) {
          LOG.warn("Multiple users have the same email adress: '" + attendee.getAddress() + "'.");
        }
      } else {
        username = listAccess.load(0, 1)[0].getUserName();
      }
    } catch (Exception e) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("exception occured while trying to get user with email " + attendee.getAddress(), e);
      }
    }
    return username;
  }

  private static void setAppointmentPriority(Appointment appointment, CalendarEvent calendarEvent) throws Exception {
    if (calendarEvent.getPriority() == null || calendarEvent.getPriority().equals(CalendarEvent.PRIORITY_NONE)
        || calendarEvent.getPriority().equals(CalendarEvent.PRIORITY_NORMAL)) {
      appointment.setImportance(Importance.Normal);
    } else if (calendarEvent.getPriority().equals(CalendarEvent.PRIORITY_LOW)) {
      appointment.setImportance(Importance.Low);
    } else if (calendarEvent.getPriority().equals(CalendarEvent.PRIORITY_HIGH)) {
      appointment.setImportance(Importance.High);
    }
  }

  private static void setEventPriority(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException {
    if (appointment.getImportance() != null) {
      switch (appointment.getImportance()) {
      case High:
        calendarEvent.setPriority(CalendarEvent.PRIORITY_HIGH);
        break;
      case Low:
        calendarEvent.setPriority(CalendarEvent.PRIORITY_LOW);
        break;
      case Normal:
        calendarEvent.setPriority(CalendarEvent.PRIORITY_NORMAL);
        break;
      }
    } else {
      calendarEvent.setPriority(CalendarEvent.PRIORITY_NORMAL);
    }
  }

  private static void setAppointmentDates(Appointment appointment, CalendarEvent calendarEvent) throws Exception {
    boolean isAllDay = isAllDayEvent(calendarEvent);
    Calendar calendar = Calendar.getInstance();

    if (isAllDay) {
      calendar.setTime(calendarEvent.getFromDateTime());
      calendar.set(Calendar.HOUR_OF_DAY, 0);
      calendar.set(Calendar.MINUTE, 0);
      calendar.set(Calendar.SECOND, 0);
      calendar.set(Calendar.MILLISECOND, TimeZone.getDefault().getRawOffset());
    } else {
      calendar = getCalendarInstance(calendarEvent.getFromDateTime());
    }
    appointment.setStart(calendar.getTime());

    if (isAllDay) {
      calendar.setTime(calendarEvent.getToDateTime());
      calendar.set(Calendar.HOUR_OF_DAY, 0);
      calendar.set(Calendar.MINUTE, 0);
      calendar.set(Calendar.SECOND, 0);
      calendar.set(Calendar.MILLISECOND, TimeZone.getDefault().getRawOffset());
    } else {
      calendar = getCalendarInstance(calendarEvent.getToDateTime());
    }
    appointment.setEnd(calendar.getTime());
    appointment.setIsAllDayEvent(isAllDay);
  }

  private static void setEventDates(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException {
    Calendar cal1 = null;
    Calendar cal2 = getCalendarInstance(appointment.getEnd());

    if (appointment.getAppointmentType().equals(AppointmentType.RecurringMaster) && appointment.getIsAllDayEvent()) {
      cal1 = getCalendarInstance(appointment.getRecurrence().getStartDate());
    } else {
      cal1 = getCalendarInstance(appointment.getStart());
    }

    if (appointment.getIsAllDayEvent()) {
      cal1.set(Calendar.HOUR_OF_DAY, 0);
      cal1.set(Calendar.MINUTE, 0);
      cal1.set(Calendar.SECOND, 0);
      cal1.set(Calendar.MILLISECOND, 0);

      cal2.add(Calendar.MILLISECOND, -1);
      cal2.set(Calendar.HOUR_OF_DAY, cal2.getActualMaximum(Calendar.HOUR_OF_DAY));
      cal2.set(Calendar.MINUTE, cal2.getActualMaximum(Calendar.MINUTE));
      cal2.set(Calendar.SECOND, cal2.getActualMaximum(Calendar.SECOND));

      calendarEvent.setFromDateTime(cal1.getTime());
      calendarEvent.setToDateTime(cal2.getTime());
    } else {
      calendarEvent.setFromDateTime(cal1.getTime());
      calendarEvent.setToDateTime(cal2.getTime());
    }

  }

  public static Calendar getCalendarInstance(Date date) {
    Calendar calendar = Calendar.getInstance();
    calendar.setTime(date);
    return calendar;
  }

  private static void setAppointmentCategory(Appointment appointment, CalendarEvent calendarEvent) throws Exception {
    if (appointment.getCategories() != null) {
      appointment.getCategories().clearList();
    }
    if (calendarEvent.getEventCategoryName() != null && !calendarEvent.getEventCategoryName().isEmpty()
        && !calendarEvent.getEventCategoryId().equals(CalendarService.DEFAULT_EVENTCATEGORY_ID_ALL)) {
      if (appointment.getCategories() == null) {
        StringList stringList = new StringList();
        appointment.setCategories(stringList);
      }
      if (!appointment.getCategories().contains(calendarEvent.getEventCategoryName())) {
        appointment.getCategories().add(calendarEvent.getEventCategoryName());
      }
    }
  }

  private static void setEventCategory(CalendarEvent calendarEvent,
                                       Appointment appointment,
                                       String username,
                                       JCRDataStorage storage) throws Exception {
    if (appointment.getCategories() != null && appointment.getCategories().getSize() > 0) {
      String categoryName = appointment.getCategories().getString(0);
      if (categoryName != null && !categoryName.isEmpty()) {
        EventCategory category = getEventCategoryByName(storage, username, getCategoryName(categoryName));
        if (category == null) {
          category = new EventCategory();
          category.setDataInit(false);
          category.setName(getCategoryName(categoryName));
          category.setId(getCategoryName(categoryName));
          storage.saveEventCategory(username, category, true);
        }
        calendarEvent.setEventCategoryId(category.getId());
        calendarEvent.setEventCategoryName(category.getName());
      }
    } else {
      EventCategory category = getEventCategoryByName(storage, username, CalendarService.DEFAULT_EVENTCATEGORY_NAME_ALL);
      if (category == null) {
        LOG.warn("Default category (" + CalendarService.DEFAULT_EVENTCATEGORY_NAME_ALL + ")of eXo Calendar is null for user: "
            + username + ".");
      } else {
        calendarEvent.setEventCategoryId(category.getId());
        calendarEvent.setEventCategoryName(category.getName());
      }
    }
  }

  private static void setAppointmentAttachements(Appointment appointment, CalendarEvent calendarEvent) throws Exception {
    List<org.exoplatform.calendar.service.Attachment> attachments = calendarEvent.getAttachment();
    AttachmentCollection attachmentCollection = appointment.getAttachments();
    if (attachments != null && !attachments.isEmpty()) {
      attachmentCollection.clear();
      for (org.exoplatform.calendar.service.Attachment attachment : attachments) {
        FileAttachment fileAttachment = attachmentCollection.addFileAttachment(attachment.getName(), attachment.getInputStream());
        fileAttachment.setContentType(attachment.getMimeType());
      }
    } else {
      attachmentCollection.clear();
      attachmentCollection.changed();
    }
  }

  private static void setEventAttachements(CalendarEvent calendarEvent, Appointment appointment) throws Exception {
    if (appointment.getHasAttachments()) {
      Appointment appointmentWithAttachment = Appointment.bind(appointment.getService(),
                                                               appointment.getId(),
                                                               new PropertySet(AppointmentSchema.Attachments));
      AttachmentCollection attachmentCollection = appointmentWithAttachment.getAttachments();
      List<org.exoplatform.calendar.service.Attachment> attachments =
                                                                    new ArrayList<org.exoplatform.calendar.service.Attachment>();
      for (microsoft.exchange.webservices.data.property.complex.Attachment attachment : attachmentCollection) {
        if (attachment instanceof FileAttachment) {
          FileAttachment fileAttachment = (FileAttachment) attachment;
          org.exoplatform.calendar.service.Attachment eXoAttachment = new org.exoplatform.calendar.service.Attachment();
          if (fileAttachment.getSize() == 0) {
            LOG.warn("Attachment '" + fileAttachment.getName() + "' of Appoitment " + appointment.getSubject() + ", start date : "
                + appointment.getStart() + " is empty, ignore it.");
            continue;
          }
          String contentType = fileAttachment.getContentType();
          if (contentType == null) {
            contentType = MimeTypes.OCTET_STREAM;
            // the mimetype of the attachment was not found
            if (LOG.isTraceEnabled()) {
              LOG.warn("No mimetype was found for attachement  '" + fileAttachment.getName() + "' of Appoitment "
                  + appointment.getSubject() + ", start date : " + appointment.getStart() + ". Use default: " + contentType);
            }
          }
          ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
          fileAttachment.load(outputStream);
          eXoAttachment.setInputStream(new ByteArrayInputStream(outputStream.toByteArray()));

          eXoAttachment.setMimeType(contentType);
          eXoAttachment.setName(fileAttachment.getName());
          eXoAttachment.setSize(fileAttachment.getSize());
          Calendar calendar = Calendar.getInstance();
          calendar.setTime(fileAttachment.getLastModifiedTime());
          eXoAttachment.setLastModified(calendar.getTimeInMillis());
          attachments.add(eXoAttachment);
        }
      }
      if (attachments.isEmpty()) {
        LOG.warn("Appointment has attachments, but it wasn't retrieved to eXo Calendar Event");
      }
      calendarEvent.setAttachment(attachments);
    } else if (calendarEvent.getAttachment() != null && !calendarEvent.getAttachment().isEmpty()) {
      calendarEvent.setAttachment(Collections.emptyList());
    }
  }

  private static void setAppointmentStatus(Appointment appointment, CalendarEvent calendarEvent) throws Exception {
    String status = (calendarEvent.getStatus() == null || calendarEvent.getStatus().isEmpty()) ? calendarEvent.getEventState()
                                                                                               : calendarEvent.getStatus();
    if (status == null) {
      status = "";
    }
    if (status.equals(CalendarEvent.ST_AVAILABLE)) {
      appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Free);
    } else if (status.equals(CalendarEvent.ST_BUSY)) {
      appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Busy);
    } else if (status.equals(CalendarEvent.ST_OUTSIDE)) {
      appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.OOF);
    } else {
      appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.NoData);
    }
  }

  private static void setEventStatus(CalendarEvent calendarEvent, Appointment appointment) throws ServiceLocalException {
    if (appointment.getLegacyFreeBusyStatus() != null) {
      switch (appointment.getLegacyFreeBusyStatus()) {
      case Free:
        calendarEvent.setStatus(CalendarEvent.ST_AVAILABLE);
        calendarEvent.setEventState(CalendarEvent.ST_AVAILABLE);
        break;
      case Busy:
        calendarEvent.setStatus(CalendarEvent.ST_BUSY);
        calendarEvent.setEventState(CalendarEvent.ST_BUSY);
        break;
      case OOF:
        calendarEvent.setStatus(CalendarEvent.ST_OUTSIDE);
        calendarEvent.setEventState(CalendarEvent.ST_OUTSIDE);
        break;
      default:
        calendarEvent.setStatus(CalendarEvent.ST_BUSY);
        calendarEvent.setEventState(CalendarEvent.ST_BUSY);
        break;
      }
    }
  }

  private static void setEventReminder(CalendarEvent event, Appointment appointment, String username) throws Exception {
    List<Reminder> reminders = new ArrayList<Reminder>();
    event.setReminders(reminders);
    try {
      if (appointment.getIsReminderSet()) {
        Reminder reminder = new Reminder();
        reminder.setFromDateTime(appointment.getReminderDueBy());
        reminder.setAlarmBefore(appointment.getReminderMinutesBeforeStart());
        reminder.setDescription("");
        reminder.setEventId(event.getId());
        reminder.setReminderType(Reminder.TYPE_POPUP);
        reminder.setReminderOwner(username);
        reminder.setRepeate(false);
        reminder.setRepeatInterval(appointment.getReminderMinutesBeforeStart());

        reminders.add(reminder);
      }
    } catch (ServiceObjectPropertyException se) {
      // occurs when no reminder set in exchange event.
      // to be checked
      // do nothing
    }
  }

  private static EventCategory getEventCategoryByName(JCRDataStorage storage,
                                                      String username,
                                                      String eventCategoryName) throws Exception {
    for (EventCategory ev : storage.getEventCategories(username)) {
      if (ev.getName().equalsIgnoreCase(eventCategoryName)) {
        return ev;
      }
    }
    return null;
  }

  private static void setApoinementSummary(Appointment appointment, CalendarEvent event) throws Exception, ServiceLocalException {
    if (event.getDescription() != null && !event.getDescription().isEmpty()) {
      appointment.setBody(MessageBody.getMessageBodyFromText(event.getDescription()));
    }
  }

  private static void setEventDescription(CalendarEvent event, Appointment appointment) throws Exception, ServiceLocalException {
    PropertySet bodyPropSet = new PropertySet(AppointmentSchema.Body);
    bodyPropSet.setRequestedBodyType(BodyType.Text);
    appointment.load(bodyPropSet);
    event.setDescription(appointment.getBody().toString());
  }

}
