using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Nop.Plugin.Shipping.CanadaPost.Models;
using Nop.Services.Configuration;
using Nop.Services.Localization;
using Nop.Services.Messages;
using Nop.Services.Security;
using Nop.Web.Framework;
using Nop.Web.Framework.Controllers;
using Nop.Web.Framework.Mvc.Filters;

namespace Nop.Plugin.Shipping.CanadaPost.Controllers
{
    [AuthorizeAdmin]
    [Area(AreaNames.Admin)]
    public class ShippingCanadaPostController : BasePluginController
    {
        #region Fields

        private readonly CanadaPostSettings _canadaPostSettings;
        private readonly ILocalizationService _localizationService;
        private readonly IPermissionService _permissionService;
        private readonly ISettingService _settingService;
        private readonly INotificationService _notificationService;

        #endregion

        #region Ctor

        public ShippingCanadaPostController(CanadaPostSettings canadaPostSettings,
                                            ILocalizationService localizationService,
                                            IPermissionService permissionService,
                                            ISettingService settingService,
                                            INotificationService notificationService)
        {
            _canadaPostSettings = canadaPostSettings;
            _localizationService = localizationService;
            _permissionService = permissionService;
            _settingService = settingService;
            _notificationService = notificationService;
        }

        #endregion

        #region Methods

        public async Task<IActionResult> Configure()
        {
            if (! await _permissionService.AuthorizeAsync(StandardPermissionProvider.ManageShippingSettings))
                return AccessDeniedView();

            var model = new CanadaPostShippingModel()
            {
                CustomerNumber = _canadaPostSettings.CustomerNumber,
                ContractId = _canadaPostSettings.ContractId,
                ApiKey = _canadaPostSettings.ApiKey,
                UseSandbox = _canadaPostSettings.UseSandbox,
                SelectedServicesCodes = _canadaPostSettings.SelectedServicesCodes
            };

            //set available services
            var availableServices = CanadaPostHelper.GetServices(null, _canadaPostSettings.ApiKey, _canadaPostSettings.UseSandbox, out string errors);
            if (availableServices != null)
            {
                model.AvailableServices = availableServices.service.Select(service => new SelectListItem
                {
                    Value = service.servicecode,
                    Text = service.servicename,
                    Selected = _canadaPostSettings.SelectedServicesCodes?.Contains(service.servicecode) ?? false
                }).ToList();
            }

            return View("~/Plugins/Shipping.CanadaPost/Views/Configure.cshtml", model);
        }

        [HttpPost]
        public async Task<IActionResult> Configure(CanadaPostShippingModel model)
        {
            if (! await _permissionService.AuthorizeAsync(StandardPermissionProvider.ManageShippingSettings))
                return AccessDeniedView();

            if (!ModelState.IsValid)
                return await Configure();

            //Canada Post page provides the API key with extra spaces
            model.ApiKey = model.ApiKey?.Replace(" : ", ":");

            //save settings
            _canadaPostSettings.CustomerNumber = model.CustomerNumber;
            _canadaPostSettings.ContractId = model.ContractId;
            _canadaPostSettings.ApiKey = model.ApiKey;
            _canadaPostSettings.UseSandbox = model.UseSandbox;
            _canadaPostSettings.SelectedServicesCodes = model.SelectedServicesCodes.ToList();
            await _settingService.SaveSettingAsync(_canadaPostSettings);

            _notificationService.SuccessNotification(await _localizationService.GetResourceAsync("Admin.Plugins.Saved"));

            return await Configure();
        }

        #endregion
    }
}
