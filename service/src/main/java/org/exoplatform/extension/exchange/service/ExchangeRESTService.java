package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import javax.annotation.security.RolesAllowed;
import javax.ws.rs.*;
import javax.ws.rs.core.*;

import org.exoplatform.common.http.HTTPStatus;
import org.exoplatform.extension.exchange.listener.IntegrationListener;
import org.exoplatform.extension.exchange.model.FolderBean;
import org.exoplatform.extension.exchange.model.UserSettings;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.rest.resource.ResourceContainer;
import org.exoplatform.services.security.ConversationState;

import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.property.complex.FolderId;

/**
 * @author Boubaker Khanfir
 */
@Path("/exchange")
public class ExchangeRESTService implements ResourceContainer, Serializable {
  private static final long serialVersionUID = -8085801604143848875L;

  private static final Log  LOG              = ExoLogger.getLogger(ExchangeRESTService.class);

  static CacheControl       cc               = new CacheControl();
  static {
    cc.setNoCache(true);
    cc.setNoStore(true);
  }

  private transient IntegrationListener integrationListener;

  private transient OrganizationService organizationService;

  public ExchangeRESTService(IntegrationListener integrationListener, OrganizationService organizationService) {
    this.integrationListener = integrationListener;
    this.organizationService = organizationService;
  }

  @GET
  @RolesAllowed("users")
  @Path("/calendars")
  @Produces({ MediaType.APPLICATION_JSON })
  public Response getCalendars() throws Exception {
    // It must be a user present in the session because of RolesAllowed
    // annotation
    String username = getCurrentUser();
    try {
      List<FolderBean> beans = new ArrayList<>();

      IntegrationService service = IntegrationService.getInstance(username);
      if (service != null) {
        List<FolderId> folderIDs = service.getAllExchangeCalendars();
        for (FolderId folderId : folderIDs) {
          Folder folder = service.getExchangeCalendar(folderId);
          if (folder != null) {
            boolean synchronizedFolder = service.isCalendarSynchronizedWithExchange(folderId.getUniqueId());
            FolderBean bean = new FolderBean(folderId.getUniqueId(), folder.getDisplayName(), synchronizedFolder);
            beans.add(bean);
          }
        }
      }
      return Response.ok(beans).cacheControl(cc).build();
    } catch (Exception e) {
      if (LOG.isTraceEnabled()) {
        LOG.trace("Exchange Inegration Service: Unable to retrieve list of calendars for user: '" + username + "'");
      }
      return Response.ok().cacheControl(cc).build();
    }
  }

  @GET
  @RolesAllowed("users")
  @Path("/syncNow")
  public Response synchronizeNow() {
    try {
      String username = getCurrentUser();
      integrationListener.synchronize(username);
      return Response.ok().build();
    } catch (Exception e) {
      LOG.error("Error while synchronizing manually the calendars", e);
      return Response.serverError().build();
    }
  }

  @GET
  @RolesAllowed("users")
  @Path("/sync")
  public Response synchronizeFolderWithExo(@QueryParam("folderId") String folderIdString) throws Exception {
    if (folderIdString == null || folderIdString.isEmpty()) {
      LOG.warn("folderId parameter is null while synchronizing.");
      return Response.noContent().build();
    }
    // It must be a user present in the session because of RolesAllowed
    // annotation
    String username = getCurrentUser();
    IntegrationService service = IntegrationService.getInstance(username);
    service.addFolderToSynchronization(folderIdString);
    integrationListener.synchronize(username);
    return Response.ok().build();
  }

  @GET
  @RolesAllowed("users")
  @Path("/unsync")
  public Response unsynchronizeFolderWithExo(@QueryParam("folderId") String folderIdString) throws Exception {
    if (folderIdString == null || folderIdString.isEmpty()) {
      LOG.warn("folderId parameter is null while unsynchronizing");
      return Response.noContent().build();
    }
    // It must be a user present in the session because of RolesAllowed
    // annotation
    String username = getCurrentUser();
    IntegrationService service = IntegrationService.getInstance(username);
    service.deleteFolderFromSynchronization(folderIdString);
    integrationListener.synchronize(username);
    return Response.ok().build();
  }

  @GET
  @RolesAllowed("users")
  @Path("/settings")
  @Produces(MediaType.APPLICATION_JSON)
  public Response getSettings() throws Exception {
    try {
      String username = getCurrentUser();

      UserSettings settings = new UserSettings();

      String exchangeServerName = IntegrationService.getUserArrtibute(organizationService,
                                                                      username,
                                                                      IntegrationService.USER_EXCHANGE_SERVER_URL_ATTRIBUTE);
      String exchangeDomainName = IntegrationService.getUserArrtibute(organizationService,
                                                                      username,
                                                                      IntegrationService.USER_EXCHANGE_SERVER_DOMAIN_ATTRIBUTE);
      String exchangeUsername = IntegrationService.getUserArrtibute(organizationService,
                                                                    username,
                                                                    IntegrationService.USER_EXCHANGE_USERNAME_ATTRIBUTE);

      settings.setServerName(exchangeServerName == null ? integrationListener.getExchangeServerURL() : exchangeServerName);
      settings.setDomainName(exchangeDomainName == null ? integrationListener.getExchangeDomain() : exchangeDomainName);
      settings.setUsername(exchangeUsername == null ? username : exchangeUsername);

      return Response.ok(settings, MediaType.APPLICATION_JSON).cacheControl(cc).build();
    } catch (Exception e) {
      return Response.status(HTTPStatus.INTERNAL_ERROR).cacheControl(cc).build();
    }
  }

  @POST
  @RolesAllowed("users")
  @Path("/settings")
  @Consumes(MediaType.APPLICATION_JSON)
  public Response setSettings(UserSettings settings) throws Exception {
    try {
      String username = getCurrentUser();

      IntegrationService.setUserArrtibute(organizationService,
                                          username,
                                          IntegrationService.USER_EXCHANGE_SERVER_URL_ATTRIBUTE,
                                          settings.getServerName());
      IntegrationService.setUserArrtibute(organizationService,
                                          username,
                                          IntegrationService.USER_EXCHANGE_SERVER_DOMAIN_ATTRIBUTE,
                                          settings.getDomainName());
      IntegrationService.setUserArrtibute(organizationService,
                                          username,
                                          IntegrationService.USER_EXCHANGE_USERNAME_ATTRIBUTE,
                                          settings.getUsername());
      IntegrationService.setUserArrtibute(organizationService,
                                          username,
                                          IntegrationService.USER_EXCHANGE_PASSWORD_ATTRIBUTE,
                                          settings.getPassword());

      integrationListener.userLoggedOut(username);
      integrationListener.startExchangeSynchronizationTask(username,
                                                           settings.getUsername(),
                                                           settings.getPassword(),
                                                           settings.getDomainName(),
                                                           settings.getServerName());

      return Response.ok().cacheControl(cc).build();
    } catch (Exception e) {
      return Response.status(HTTPStatus.INTERNAL_ERROR).cacheControl(cc).build();
    }
  }

  private String getCurrentUser() {
    return ConversationState.getCurrent().getIdentity().getUserId();
  }

}
