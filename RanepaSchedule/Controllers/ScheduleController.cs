using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using RanepaSchedule.Services;

namespace RanepaSchedule.Controllers
{
    public class ScheduleController
    {
        private readonly ScheduleService _scheduleService;

        public ScheduleController(ScheduleService scheduleService)
        {
            _scheduleService = scheduleService;
        }

        [HttpGet]
        [Route("")]
        public async Task<IActionResult> GetSchedule()
        {
            return new JsonResult(await _scheduleService.GetSchedule());
        }
    }
}