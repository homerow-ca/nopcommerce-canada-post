using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Nop.Core;
using Nop.Core.Domain.Shipping;
using Nop.Services.Configuration;
using Nop.Services.Directory;
using Nop.Services.Installation;
using Nop.Services.Localization;
using Nop.Services.Logging;
using Nop.Services.Plugins;
using Nop.Services.Shipping;
using Nop.Services.Shipping.Tracking;

namespace Nop.Plugin.Shipping.CanadaPost
{
    /// <summary>
    /// Canada post computation method
    /// </summary>
    public class CanadaPostComputationMethod : BasePlugin, IShippingRateComputationMethod
    {
        #region Fields

        private readonly CanadaPostSettings _canadaPostSettings;
        private readonly ICurrencyService _currencyService;
        private readonly ILocalizationService _localizationService;
        private readonly ILogger _logger;
        private readonly IMeasureService _measureService;
        private readonly ISettingService _settingService;
        private readonly IShippingService _shippingService;
        private readonly IWebHelper _webHelper;
        private readonly ICountryService _countryService;

        #endregion

        #region Ctor

        public CanadaPostComputationMethod(CanadaPostSettings canadaPostSettings,
                                           ICurrencyService currencyService,
                                           ILocalizationService localizationService,
                                           ILogger logger,
                                           IMeasureService measureService,
                                           ISettingService settingService,
                                           IShippingService shippingService,
                                           IWebHelper webHelper,
                                           ICountryService countryService)
        {
            _canadaPostSettings = canadaPostSettings;
            _currencyService = currencyService;
            _localizationService = localizationService;
            _logger = logger;
            _measureService = measureService;
            _settingService = settingService;
            _shippingService = shippingService;
            _webHelper = webHelper;
            _countryService = countryService;
        }

        #endregion

        #region Utilities

        /// <summary>
        /// Calculate parcel weight in kgs
        /// </summary>
        /// <param name="getShippingOptionRequest">A request for getting shipping options</param>
        /// <param name="weight">Weight</param>
        private async Task<decimal> GetWeight(GetShippingOptionRequest getShippingOptionRequest)
        {
            var usedMeasureWeight = await _measureService.GetMeasureWeightBySystemKeywordAsync("kg");
            if (usedMeasureWeight == null)
                throw new NopException("CanadaPost shipping service. Could not load \"kg\" measure weight");

            var weight = await _shippingService.GetTotalWeightAsync(getShippingOptionRequest, ignoreFreeShippedItems: true);
            return await _measureService.ConvertFromPrimaryMeasureWeightAsync(weight, usedMeasureWeight);
        }

        /// <summary>
        /// Calculate parcel length, width, height in centimeters
        /// </summary>
        /// <param name="getShippingOptionRequest">A request for getting shipping options</param>
        /// <param name="length">Length</param>
        /// <param name="width">Width</param>
        /// <param name="height">height</param>
        private async Task<(decimal length, decimal width, decimal height)> GetDimensions(GetShippingOptionRequest getShippingOptionRequest)
        {
            var usedMeasureDimension = await _measureService.GetMeasureDimensionBySystemKeywordAsync("meters");
            if (usedMeasureDimension == null)
                throw new NopException("CanadaPost shipping service. Could not load \"meter(s)\" measure dimension");

            var (width, length, height) = await _shippingService.GetDimensionsAsync(getShippingOptionRequest.Items, true);

            //In the Canada Post API length is longest dimension, width is second longest dimension, height is shortest dimension
            var dimensions = new List<decimal> { length, width, height };
            dimensions.Sort();
            return (
                Math.Round(await _measureService.ConvertFromPrimaryMeasureDimensionAsync(dimensions[2], usedMeasureDimension) * 100, 1),
                Math.Round(await _measureService.ConvertFromPrimaryMeasureDimensionAsync(dimensions[1], usedMeasureDimension) * 100, 1),
                Math.Round(await _measureService.ConvertFromPrimaryMeasureDimensionAsync(dimensions[0], usedMeasureDimension) * 100, 1)
            );
        }

        /// <summary>
        /// Get shipping price in the primary store currency
        /// </summary>
        /// <param name="price">Price in CAD currency</param>
        /// <returns>Price amount</returns>
        private async Task<decimal> PriceToPrimaryStoreCurrency(decimal price)
        {
            var cad = await _currencyService.GetCurrencyByCodeAsync("CAD");
            if (cad == null)
                throw new Exception("CAD currency cannot be loaded");

            return await _currencyService.ConvertToPrimaryStoreCurrencyAsync(price, cad);
        }
        #endregion

        #region Methods

        /// <summary>
        ///  Gets available shipping options
        /// </summary>
        /// <param name="getShippingOptionRequest">A request for getting shipping options</param>
        /// <returns>Represents a response of getting shipping rate options</returns>
        public async Task<GetShippingOptionResponse> GetShippingOptionsAsync(GetShippingOptionRequest getShippingOptionRequest)
        {
            if (getShippingOptionRequest == null)
                throw new ArgumentNullException(nameof(getShippingOptionRequest));

            if (getShippingOptionRequest.Items == null)
                return new GetShippingOptionResponse { Errors = new List<string> { "No shipment items" } };

            if (getShippingOptionRequest.ShippingAddress == null)
                return new GetShippingOptionResponse { Errors = new List<string> { "Shipping address is not set" } };

            if (getShippingOptionRequest.ShippingAddress.CountryId == null)
                return new GetShippingOptionResponse { Errors = new List<string> { "Shipping country is not set" } };

            if (string.IsNullOrEmpty(getShippingOptionRequest.ZipPostalCodeFrom))
                return new GetShippingOptionResponse { Errors = new List<string> { "Origin postal code is not set" } };

            var country = await _countryService.GetCountryByIdAsync((int) getShippingOptionRequest.ShippingAddress.CountryId);

            //get available services
            var availableServices = CanadaPostHelper.GetServices(country.TwoLetterIsoCode,
                _canadaPostSettings.ApiKey, _canadaPostSettings.UseSandbox, out string errors);
            if (availableServices == null)
                return new GetShippingOptionResponse { Errors = new List<string> { errors } };

            //create object for the get rates requests
            var result = new GetShippingOptionResponse();
            object destinationCountry;
            switch (country.TwoLetterIsoCode.ToLowerInvariant())
            {
                case "us":
                    destinationCountry = new mailingscenarioDestinationUnitedstates
                    {
                        zipcode = getShippingOptionRequest.ShippingAddress.ZipPostalCode
                    };
                    break;
                case "ca":
                    destinationCountry = new mailingscenarioDestinationDomestic
                    {
                        postalcode = getShippingOptionRequest.ShippingAddress.ZipPostalCode.Replace(" ", string.Empty).ToUpperInvariant()
                    };
                    break;
                default:
                    destinationCountry = new mailingscenarioDestinationInternational
                    {
                        countrycode = country.TwoLetterIsoCode
                    };
                    break;
            }

            var mailingScenario = new mailingscenario
            {
                quotetype = mailingscenarioQuotetype.counter,
                quotetypeSpecified = true,
                originpostalcode = getShippingOptionRequest.ZipPostalCodeFrom.Replace(" ", string.Empty).ToUpperInvariant(),
                destination = new mailingscenarioDestination
                {
                    Item = destinationCountry
                }
            };

            //set contract customer properties
            if (!string.IsNullOrEmpty(_canadaPostSettings.CustomerNumber))
            {
                mailingScenario.quotetype = mailingscenarioQuotetype.commercial;
                mailingScenario.customernumber = _canadaPostSettings.CustomerNumber;
                mailingScenario.contractid = !string.IsNullOrEmpty(_canadaPostSettings.ContractId) ? _canadaPostSettings.ContractId : null;
            }

            //get original parcel characteristics
            var originalWeight = await GetWeight(getShippingOptionRequest);
            var (originalLength, originalWidth, originalHeight) = await GetDimensions(getShippingOptionRequest);

            //get rate for selected services
            var errorSummary = new StringBuilder();
            var selectedServices = availableServices.service.Where(service => _canadaPostSettings.SelectedServicesCodes.Contains(service.servicecode));
            foreach (var service in selectedServices)
            {
                var currentService = CanadaPostHelper.GetServiceDetails(_canadaPostSettings.ApiKey, service.link.href, service.link.mediatype, out errors);
                if (currentService != null)
                {
                    #region parcels count calculation

                    var totalParcels = 1;
                    var restrictions = currentService.restrictions;

                    //parcels count by weight
                    var maxWeight = restrictions?.weightrestriction?.maxSpecified ?? false 
                        ? restrictions.weightrestriction.max : int.MaxValue;
                    if (originalWeight * 1000 > maxWeight)
                    {
                        var parcelsOnWeight = Convert.ToInt32(Math.Ceiling(originalWeight * 1000 / maxWeight));
                        if (parcelsOnWeight > totalParcels)
                            totalParcels = parcelsOnWeight;
                    }

                    //parcels count by length
                    var maxLength = restrictions?.dimensionalrestrictions?.length?.maxSpecified ?? false 
                        ? restrictions.dimensionalrestrictions.length.max : int.MaxValue;
                    if (originalLength > maxLength)
                    {
                        var parcelsOnLength = Convert.ToInt32(Math.Ceiling(originalLength / maxLength));
                        if (parcelsOnLength > totalParcels)
                            totalParcels = parcelsOnLength;
                    }

                    //parcels count by width
                    var maxWidth = restrictions?.dimensionalrestrictions?.width?.maxSpecified ?? false 
                        ? restrictions.dimensionalrestrictions.width.max : int.MaxValue;
                    if (originalWidth > maxWidth)
                    {
                        var parcelsOnWidth = Convert.ToInt32(Math.Ceiling(originalWidth / maxWidth));
                        if (parcelsOnWidth > totalParcels)
                            totalParcels = parcelsOnWidth;
                    }

                    //parcels count by height
                    var maxHeight = restrictions?.dimensionalrestrictions?.height?.maxSpecified ?? false 
                        ? restrictions.dimensionalrestrictions.height.max : int.MaxValue;
                    if (originalHeight > maxHeight)
                    {
                        var parcelsOnHeight = Convert.ToInt32(Math.Ceiling(originalHeight / maxHeight));
                        if (parcelsOnHeight > totalParcels)
                            totalParcels = parcelsOnHeight;
                    }

                    //parcel count by girth
                    var lengthPlusGirthMax = restrictions?.dimensionalrestrictions?.lengthplusgirthmaxSpecified ?? false 
                        ? restrictions.dimensionalrestrictions.lengthplusgirthmax : int.MaxValue;
                    var lengthPlusGirth = 2 * (originalWidth + originalHeight) + originalLength;
                    if (lengthPlusGirth > lengthPlusGirthMax)
                    {
                        var parcelsOnHeight = Convert.ToInt32(Math.Ceiling(lengthPlusGirth / lengthPlusGirthMax));
                        if (parcelsOnHeight > totalParcels)
                            totalParcels = parcelsOnHeight;
                    }

                    //parcel count by sum of length, width and height
                    var lengthWidthHeightMax = restrictions?.dimensionalrestrictions?.lengthheightwidthsummaxSpecified ?? false 
                        ? restrictions.dimensionalrestrictions.lengthheightwidthsummax : int.MaxValue;
                    var lengthWidthHeight = originalLength + originalWidth + originalHeight;
                    if (lengthWidthHeight > lengthWidthHeightMax)
                    {
                        var parcelsOnHeight = Convert.ToInt32(Math.Ceiling(lengthWidthHeight / lengthWidthHeightMax));
                        if (parcelsOnHeight > totalParcels)
                            totalParcels = parcelsOnHeight;
                    }

                    #endregion

                    //set parcel characteristics
                    mailingScenario.services = new[] { currentService.servicecode };
                    mailingScenario.parcelcharacteristics = new mailingscenarioParcelcharacteristics
                    {
                        weight = Math.Round(originalWeight / totalParcels, 3),
                        dimensions = new mailingscenarioParcelcharacteristicsDimensions
                        {
                            length = Math.Round(originalLength / totalParcels, 1),
                            width = Math.Round(originalWidth / totalParcels, 1),
                            height = Math.Round(originalHeight / totalParcels, 1)
                        }
                    };

                    //get rate
                    var priceQuotes = CanadaPostHelper.GetShippingRates(mailingScenario, _canadaPostSettings.ApiKey, _canadaPostSettings.UseSandbox, out errors);
                    if (priceQuotes != null)
                        foreach (var option in priceQuotes.pricequote)
                        {
                            var shippingOption = new ShippingOption
                            {
                                Name = option.servicename,
                                Rate = await PriceToPrimaryStoreCurrency(option.pricedetails.due * totalParcels)
                            };
                            if (!string.IsNullOrEmpty(option.servicestandard?.expectedtransittime))
                                shippingOption.Description = $"Delivery in {option.servicestandard.expectedtransittime} days {(totalParcels > 1 ? $"into {totalParcels} parcels" : string.Empty)}";
                            result.ShippingOptions.Add(shippingOption);
                        }
                    else
                        errorSummary.AppendLine(errors);
                }
                else
                    errorSummary.AppendLine(errors);
            }

            //write errors
            var errorString = errorSummary.ToString();
            if (!string.IsNullOrEmpty(errorString))
                await _logger.ErrorAsync(errorString);
            if (!result.ShippingOptions.Any())
                result.AddError(errorString);

            return result;
        }

        /// <summary>
        /// Gets fixed shipping rate (if shipping rate computation method allows it and the rate can be calculated before checkout).
        /// </summary>
        /// <param name="getShippingOptionRequest">A request for getting shipping options</param>
        /// <returns>Fixed shipping rate; or null in case there's no fixed shipping rate</returns>
        public async Task<decimal?> GetFixedRateAsync(GetShippingOptionRequest getShippingOptionRequest)
        {
            return null;
        }

        public async Task<IShipmentTracker> GetShipmentTrackerAsync()
        {
            return new CanadaPostShipmentTracker(_canadaPostSettings, _logger); 
        }

        /// <summary>
        /// Gets a configuration page URL
        /// </summary>
        public override string GetConfigurationPageUrl()
        {
            return $"{_webHelper.GetStoreLocation()}Admin/ShippingCanadaPost/Configure";
        }

        /// <summary>
        /// Install the plugin
        /// </summary>
        public override async Task InstallAsync()
        {
            //settings
            var settings = new CanadaPostSettings
            {
                 UseSandbox = true
            };
            await _settingService.SaveSettingAsync(settings);

            //locales
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Api", "API key");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Api.Hint", "Specify Canada Post API key.");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.ContractId", "Contract ID");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.ContractId.Hint", "Specify contract identifier.");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.CustomerNumber", "Customer number");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.CustomerNumber.Hint", "Specify customer number.");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Services", "Available services");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Services.Hint", "Select the services you want to offer to customers.");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.UseSandbox", "Use Sandbox");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.UseSandbox.Hint", "Check to enable Sandbox (testing environment).");
            await _localizationService.AddOrUpdateLocaleResourceAsync("Plugins.Shipping.CanadaPost.Instructions", "<p>To configure plugin follow one of these steps:<br />1. If you are a Canada Post commercial customer, fill Customer number, Contract ID and API key below.<br />2. If you are a Solutions for Small Business customer, specify your Customer number and API key below.<br />3. If you are a non-contracted customer or you want to use the regular price of shipping paid by customers, fill the API key field only.<br /><br /><em>Note: Canada Post gateway returns shipping price in the CAD currency, ensure that you have correctly configured exchange rate from PrimaryStoreCurrency to CAD.</em></p>");

            await base.InstallAsync();
        }
        
        /// <summary>
        /// Uninstall the plugin
        /// </summary>
        public override async Task UninstallAsync()
        {
            //settings
            await _settingService.DeleteSettingAsync<CanadaPostSettings>();

            //locales
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Api");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Api.Hint");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.ContractId");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.ContractId.Hint");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.CustomerNumber");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.CustomerNumber.Hint");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Services");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.Services.Hint");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.UseSandbox");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Fields.UseSandbox.Hint");
            await _localizationService.DeleteLocaleResourceAsync("Plugins.Shipping.CanadaPost.Instructions");

            await base.UninstallAsync();
        }

        #endregion
    }
}
