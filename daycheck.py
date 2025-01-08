import datetime
from enum import Enum

class Holiday(Enum):
    def __new__(cls, english, chinese, days):
        obj = object.__new__(cls)
        obj._value_ = english  # 使用英文作为Enum的值
        obj.chinese = chinese  # 中文名称
        obj.days = days        # 放假天数
        return obj

    new_years_day = "New Year's Day", "元旦", 1
    spring_festival = "Spring Festival", "春节", 4
    tomb_sweeping_day = "Tomb-sweeping Day", "清明", 1
    labour_day = "Labour Day", "劳动节", 2
    dragon_boat_festival = "Dragon Boat Festival", "端午", 1
    national_day = "National Day", "国庆节", 3
    mid_autumn_festival = "Mid-autumn Festival", "中秋", 1

    # 特殊节假日
    anti_fascist_70th_day = "Anti-Fascist 70th Day", "中国人民抗日战争暨世界反法西斯战争胜利70周年纪念日", 1

class DayCheck:
    def __init__(self, year):
        self.year = year        
        self.holidays = {
            datetime.date(year=2025, month=1, day=1): Holiday.new_years_day.chinese,
            datetime.date(year=2025, month=1, day=28): Holiday.spring_festival.chinese,
            datetime.date(year=2025, month=1, day=29): Holiday.spring_festival.chinese,
            datetime.date(year=2025, month=1, day=30): Holiday.spring_festival.chinese,
            datetime.date(year=2025, month=1, day=31): Holiday.spring_festival.chinese,
            datetime.date(year=2025, month=4, day=4): Holiday.tomb_sweeping_day.chinese,
            datetime.date(year=2025, month=5, day=1): Holiday.labour_day.chinese,
            datetime.date(year=2025, month=5, day=2): Holiday.labour_day.chinese,
            datetime.date(year=2025, month=5, day=31): Holiday.dragon_boat_festival.chinese,
            datetime.date(year=2025, month=10, day=1): Holiday.national_day.chinese,
            datetime.date(year=2025, month=10, day=2): Holiday.national_day.chinese,
            datetime.date(year=2025, month=10, day=3): Holiday.national_day.chinese,
            datetime.date(year=2025, month=10, day=6): Holiday.mid_autumn_festival.chinese,
        }
    
        self.restdays = set()  # 休息日集合
        self.workdays = set()  # 工作日集合
        self.init_days()

    def is_leap_year(self, year):
        """判断是否为闰年"""
        return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)

    def get_day_type(self, date: datetime.date):
        """判断日期所属的类型，返回 holiday/workday/restday"""
        if self.is_holiday(date):
            return "holiday"
        elif self.is_restday(date):
            return "restday"
        elif self.is_workday(date):
            return "workday"
        else:
            return "unknown"

    def init_days(self):
        """初始化工作日和休息日数据，确保一年中的每一天都被涵盖"""
        days_in_year = 366 if self.is_leap_year(self.year) else 365

        for day_offset in range(days_in_year):  # 一年365或366天
            current_date = datetime.date(year=self.year, month=1, day=1) + datetime.timedelta(days=day_offset)
            
            # 如果该日期是节假日，则加入holidays
            if current_date in self.holidays:
                continue  # 跳过节假日
            
            # 判断是否为周末（周六、周日）
            if current_date.weekday() >= 5:  # 周六或周日
                self.restdays.add(current_date)
            else:
                self.workdays.add(current_date)

        # 此后新增调休的信息
        self.set_restdays_mult("2025-1-26","2025-2-4")
        self.set_workday("2025-1-25")
        self.set_workday("2025-2-8")

    def is_holiday(self, date: datetime.date):
        """判断是否是节假日"""
        return date in self.holidays

    def is_workday(self, date: datetime.date):
        """判断是否是工作日"""
        return date in self.workdays

    def is_restday(self, date: datetime.date):
        """判断是否是休息日"""
        return date in self.restdays

    def set_day_type(self, date: datetime.date, day_type: str):
        """设置日期的类型为休息日或工作日"""
        if self.is_holiday(date):
            # print(f"该日期 {date} 是假期，不能设置为 {day_type}")
            return

        if day_type == "restday":
            if self.is_restday(date):
                # print(f"该日期 {date} 已是休息日")
                return
            if self.is_workday(date):
                self.workdays.remove(date)
            self.restdays.add(date)

        elif day_type == "workday":
            if self.is_restday(date):
                self.restdays.remove(date)
            self.workdays.add(date)

    def set_restdays(self, date):
        """设置单个日期为休息日"""
        date = self._normalize_date_format(date)
        self.set_day_type(date, "restday")

    def set_restdays_mult(self, start_date, end_date):
        """设置一段日期为休息日"""
        start_date = self._normalize_date_format(start_date)
        end_date = self._normalize_date_format(end_date)
        
        current_date = start_date
        while current_date <= end_date:
            if self.is_holiday(current_date):
                # print(f"该日期 {current_date} 是假期，跳过")
                pass
            else:
                self.set_day_type(current_date, "restday")
            current_date += datetime.timedelta(days=1)

    def set_workday(self, date):
        """设置单个日期为工作日"""
        date = self._normalize_date_format(date)
        self.set_day_type(date, "workday")

    def set_workday_mult(self, start_date, end_date):
        """设置一段日期为工作日"""
        start_date = self._normalize_date_format(start_date)
        end_date = self._normalize_date_format(end_date)

        current_date = start_date
        while current_date <= end_date:
            if self.is_holiday(current_date):
                # print(f"该日期 {current_date} 是假期，跳过")
                pass
            else:
                self.set_day_type(current_date, "workday")
            current_date += datetime.timedelta(days=1)

    def _normalize_date_format(self, date_str):
        """将日期字符串规范化为 YYYY-MM-DD 格式，确保月份和日期为两位数"""
        if isinstance(date_str, str):
            parts = date_str.split("-")
            year = parts[0]
            month = parts[1].zfill(2)  # 补齐月份为两位数
            day = parts[2].zfill(2)    # 补齐日期为两位数
            return datetime.datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d").date()
        elif isinstance(date_str, datetime.date):
            return date_str
        else:
            raise ValueError("日期格式无效，必须为字符串或 datetime.date 对象")

    