class Technician:
    def __init__(self, name, job='NA', job_type='NA', spqh=0, mobile_support_hours=0, mac_support_hours=0,
                 iphone_repair_hours=0, mac_repair_hours=0, repair_pickup_hours=0,
                 gb_on_point_hours=0, daily_download_hours=0, guided_hours=0,
                 connection_hours=0, customers_helped=0,
                 mac_queue_sessions_delivered=0, mac_queue_avg_session_duration=0,
                 mobile_queue_sessions_delivered=0, mobile_queue_avg_session_duration=0, total_hours=0, nps=0,
                 tms=0, sur=0, business_intro=0, apple_care=0, trade_in=0, year=0, month=0, survey=0, full_survey=0,
                 opportunities=0):
        self.name = name
        self.job = job
        self.job_type = job_type
        self.spqh = spqh
        self.mobile_support_hours = mobile_support_hours
        self.mac_support_hours = mac_support_hours
        self.iphone_repair_hours = iphone_repair_hours
        self.mac_repair_hours = mac_repair_hours
        self.repair_pickup_hours = repair_pickup_hours
        self.gb_on_point_hours = gb_on_point_hours
        self.daily_download_hours = daily_download_hours
        self.guided_hours = guided_hours
        self.connection_hours = connection_hours
        self.customers_helped = customers_helped
        self.mac_queue_sessions_delivered = mac_queue_sessions_delivered
        self.mac_queue_avg_session_duration = mac_queue_avg_session_duration
        self.mobile_queue_sessions_delivered = mobile_queue_sessions_delivered
        self.mobile_queue_avg_session_duration = mobile_queue_avg_session_duration
        self.total_hours = total_hours
        self.nps = nps
        self.tms = tms
        self.sur = sur
        self.business_intro = business_intro
        self.apple_care = apple_care
        self.trade_in = trade_in
        self.year = year
        self.month = month
        self.survey = survey
        self.full_survey = full_survey
        self.opportunities = opportunities

    def display_info(self):
        # Display information
        print("Name:", self.name)
        print("Job:", self.job)
        print("Type:", self.job_type)
        print("SPQH:", self.spqh)
        print("Mobile Support hours:", self.mobile_support_hours)
        print("Mac Support hours:", self.mac_support_hours)
        print("iPhone Repair hours:", self.iphone_repair_hours)
        print("Mac Repair hours:", self.mac_repair_hours)
        print("Repair Pickup hours:", self.repair_pickup_hours)
        print("GB On Point hours:", self.gb_on_point_hours)
        print("Daily Download hours:", self.daily_download_hours)
        print("Guided hours:", self.guided_hours)
        print("Connection hours:", self.connection_hours)
        print("Total hours:", self.total_hours)
        print("Customers Helped:", self.customers_helped)
        print("Mac Queue Sessions Delivered:", self.mac_queue_sessions_delivered)
        print("Mac Queue Avg Session Duration:", self.mac_queue_avg_session_duration)
        print("Mobile Queue Sessions Delivered:", self.mobile_queue_sessions_delivered)
        print("Mobile Queue Avg Session Duration:", self.mobile_queue_avg_session_duration)
