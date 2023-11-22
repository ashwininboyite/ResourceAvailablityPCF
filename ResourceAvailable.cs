using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;
using Newtonsoft.Json;

namespace ACC.CE.AvailableResourcePlugin
{
    public class ResourceAvailable : IPlugin
    {
        private class FinalTimeSlotDetails
        {
            public DateTime StartTime { get; set; }
            public DateTime EndTime { get; set; }
            public string ResourceName { get; set; }
            public string ResourceId { get; set; }
            public string ResourceType { get; set; }
        }

        public class BookableResource
        {
            public int key { get; set; }
            public string resourceName { get; set; }
            public DateTime startTime { get; set; }
            public DateTime endTime { get; set; }
            public Guid resourceId { get; set; }
            public string resourceType { get; set; }
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            //Initializing Service Context.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference targetEntity = (EntityReference)context.InputParameters["Target"];
                string operationType = context.InputParameters["OperationType"].ToString();
                Guid workorderId = targetEntity.Id;
                string timeSlots = string.Empty;
                int? timeZoneCode = RetrieveCurrentUsersSettings(service, context);

                if (operationType == "RetriveTimeSlot")
                {
                    Entity workOrderDetails = RetriveWorkOrderDetails(workorderId, service);

                    DateTime fromDate = workOrderDetails.Contains("acc_dispatchdatetime") ? workOrderDetails.GetAttributeValue<DateTime>("acc_dispatchdatetime") : DateTime.MinValue;

                    if (fromDate != DateTime.MinValue)
                    {
                        DateTime toDate = fromDate.AddHours(8);

                        EntityCollection timeSlotCollection = FetchTimeSlots(fromDate, toDate, service);

                        if (timeSlotCollection != null && timeSlotCollection.Entities != null && timeSlotCollection.Entities.Count > 0)
                        {
                            timeSlots = SetTimeSlotOutput(timeSlotCollection, fromDate, toDate, timeZoneCode, service);
                        }
                    }
                    context.OutputParameters["ResourceAvailabilty"] = timeSlots;
                }
                else if (operationType == "BookableResourceBooking")
                {
                    string selectedResources = (string)context.InputParameters["SelectedResources"];
                    string jsonString = selectedResources.Trim('"').Replace("\\\"", "\"");
                    var bookableResoureBooking = JsonConvert.DeserializeObject<BookableResource>(jsonString);

                    if (bookableResoureBooking != null)
                    {
                        DateTime parseddStartDateTime = DateTime.ParseExact(bookableResoureBooking.startTime.ToString("yyyy-MM-dd HH:mm:ss"), "yyyy-MM-dd HH:mm:ss", null);
                        DateTime utcStartTime = RetrieveUTCTimeFromLocalTime(parseddStartDateTime, timeZoneCode, service);

                        DateTime parseddEndDateTime = DateTime.ParseExact(bookableResoureBooking.endTime.ToString("yyyy-MM-dd HH:mm:ss"), "yyyy-MM-dd HH:mm:ss", null);
                        DateTime utcEndTime = RetrieveUTCTimeFromLocalTime(parseddEndDateTime, timeZoneCode, service);

                        CreateBookableResourceBooking(workorderId, utcStartTime, utcEndTime, bookableResoureBooking.resourceId, service);
                    }
                    context.OutputParameters["ResourceAvailabilty"] = "record created successfully";
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void CreateBookableResourceBooking(Guid workOrderId, DateTime startTime, DateTime endTime, Guid resourceId, IOrganizationService service)
        {
            try
            {
                Entity bookableResourceBooking = new Entity("bookableresourcebooking");
                bookableResourceBooking["msdyn_workorder"] = new EntityReference("msdyn_workorder", workOrderId);
                bookableResourceBooking["starttime"] = startTime;
                bookableResourceBooking["endtime"] = endTime;
                bookableResourceBooking["resource"] = new EntityReference("bookableresource", resourceId);
                bookableResourceBooking["bookingstatus"] = new EntityReference("bookingstatus", new Guid("74d847d0-5c88-ee11-8178-001dd806e56f"));
                bookableResourceBooking["msdyn_actualarrivaltime"] = startTime;
                service.Create(bookableResourceBooking);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private string SetTimeSlotOutput(EntityCollection timeSlotCollection, DateTime fromDate, DateTime toDate, int? timeZoneCode, IOrganizationService service)
        {
            string timeSlots = string.Empty;
            List<FinalTimeSlotDetails> finalTimeSlots = new List<FinalTimeSlotDetails>();

            try
            {
                foreach (Entity timeSlot in timeSlotCollection.Entities)
                {
                    bool potential = false;
                    Guid resourceId = Guid.Empty;
                    DateTime startTime, endTime, arrivalTime;
                    Guid bookableResourceBookId = Guid.Empty;
                    string resourceName = string.Empty;

                    potential = timeSlot.GetAttributeValue<Boolean>("Potential");
                    startTime = timeSlot.GetAttributeValue<DateTime>("StartTime");
                    endTime = timeSlot.GetAttributeValue<DateTime>("EndTime");

                    DateTime fromDateLocal = RetrieveLocalTimeFromUTCTime(fromDate, timeZoneCode, service);
                    DateTime toDateLocal = fromDateLocal.AddHours(8);

                    if (potential && startTime == fromDate && endTime == toDate)
                    {
                        resourceId = timeSlot.GetAttributeValue<Entity>("Resource").GetAttributeValue<EntityReference>("Resource").Id;
                        resourceName = timeSlot.GetAttributeValue<Entity>("Resource").GetAttributeValue<EntityReference>("Resource").Name;
                        startTime = timeSlot.GetAttributeValue<DateTime>("StartTime");
                        endTime = timeSlot.GetAttributeValue<DateTime>("EndTime");
                        arrivalTime = timeSlot.GetAttributeValue<DateTime>("ArrivalTime");

                        FinalTimeSlotDetails finalTimeSlotDetails = new FinalTimeSlotDetails();
                        finalTimeSlotDetails.StartTime = fromDateLocal;
                        finalTimeSlotDetails.EndTime = toDateLocal;
                        finalTimeSlotDetails.ResourceName = timeSlot.GetAttributeValue<Entity>("Resource").GetAttributeValue<EntityReference>("Resource").Name;
                        finalTimeSlotDetails.ResourceId = timeSlot.GetAttributeValue<Entity>("Resource").GetAttributeValue<EntityReference>("Resource").Id.ToString();
                        finalTimeSlotDetails.ResourceType = "User";

                        finalTimeSlots.Add(finalTimeSlotDetails);
                    }

                }
                return timeSlots = System.Text.Json.JsonSerializer.Serialize(finalTimeSlots);
                //return timeSlots = JsonConvert.SerializeObject(finalTimeSlots);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private Entity RetriveWorkOrderDetails(Guid workOrderId, IOrganizationService service)
        {
            try
            {
                Entity workOrder = service.Retrieve("msdyn_workorder", workOrderId, new ColumnSet("acc_dispatchdatetime"));
                return workOrder;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private int? RetrieveCurrentUsersSettings(IOrganizationService service, IPluginExecutionContext context)
        {
            var currentUserSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal, context.UserId)
                        }
                    }
                }).Entities[0].ToEntity<Entity>();
            //return time zone code
            return (int?)currentUserSettings.Attributes["timezonecode"];
        }

        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, int? timeZoneCode, IOrganizationService service)
        {
            if (!timeZoneCode.HasValue)
                return DateTime.Now;
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime
            };
            var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
            return response.LocalTime;
        }

        private DateTime RetrieveUTCTimeFromLocalTime(DateTime localTime, int? timeZoneCode, IOrganizationService service)
        {
            if (!timeZoneCode.HasValue)
                return DateTime.Now;
            var request = new UtcTimeFromLocalTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                LocalTime = localTime
            };
            var response = (UtcTimeFromLocalTimeResponse)service.Execute(request);
            return response.UtcTime;
        }

        private EntityCollection FetchTimeSlots(DateTime fromDate, DateTime toDate, IOrganizationService service)
        {
            EntityCollection timeSlots = null;
            try
            {
                Entity settings = new Entity("organization");
                settings["UseRealTimeResourceLocation"] = true;
                settings["ConsiderTravelTime"] = false;
                settings["ConsiderSlotsWithOverlappingBooking"] = false;
                settings["ConsiderSlotsWithLessThanRequiredDuration"] = false;
                settings["ConsiderSlotsWithProposedBookings"] = false;
                Entity requirement = new Entity("msdyn_resourcerequirement");
                requirement["msdyn_fromdate"] = fromDate;
                requirement["msdyn_todate"] = toDate;

                var entityCollectionResourceType = new EntityCollection();
                entityCollectionResourceType.Entities.Add(new Entity
                {
                    Id = new Guid(),
                    LogicalName = "ResourceTypes",
                    Attributes =
                new AttributeCollection
                {new KeyValuePair<string, object>("value", 3)}
                });

    //            var entityCollectionCharacteristics = new EntityCollection();
    //            entityCollectionCharacteristics.Entities.Add(new Entity
    //            {
    //                Id = new Guid(),
    //                LogicalName = "bookableresourcecharacteristic",
    //                Attributes = new AttributeCollection
    //{
    //    new KeyValuePair<string, object>("bookableresourcecharacteristicid", "46b0854a-7288-ee11-8178-001dd806ee0c"),
    //}
    //            });

                var entityCollectionCharacteristics = new EntityCollection();
                entityCollectionCharacteristics.Entities.Add(new Entity
                {
                    Id = new Guid(),
                    LogicalName = "bookableresourcecharacteristic",
                    Attributes =
                new AttributeCollection
                {
new KeyValuePair<string, object>("characteristic", "124a6764-4821-ee11-8f6c-001dd809b6a9"),
new KeyValuePair<string, object>("ratingvalue", "2")
                }
                });
                var constraints = new Entity()
                {
                    LogicalName = "organization",
                    Id = new Guid(),
                    Attributes = new AttributeCollection
{
new KeyValuePair<string, object>("Characteristics", entityCollectionCharacteristics)
}
                }; //Characteristics

                //var characteristicsAndRatingConstraints = new Entity("organization");
                //characteristicsAndRatingConstraints["Characteristics"] = entityCollectionCharacteristics;

                Entity resourceSpecification = new Entity("organization");
                resourceSpecification["ResourceTypes"] = entityCollectionResourceType;
                //resourceSpecification["CharacteristicsAndRatingConstraints"] = characteristicsAndRatingConstraints;
                resourceSpecification["Constraints"] = constraints;

                var response = service.Execute(
                new OrganizationRequest("msdyn_SearchResourceAvailability")//""msdyn_getresourceavailability")
                {
                    Parameters = {
                        { "Version", "3.0" },
                        { "Requirement", requirement},
                        { "Settings", settings},
                        {"ResourceSpecification", resourceSpecification}
                }
                });

                timeSlots = (EntityCollection)response.Results["TimeSlots"];

                return timeSlots;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
