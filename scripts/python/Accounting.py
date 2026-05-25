# coding: utf-8

from __future__ import unicode_literals
import calendar
from datetime import date, datetime, timedelta
from decimal import Decimal
from enum import Enum, auto
import numbers
import json
import time
import traceback

import uno

from com.sun.star.table import CellRangeAddress
from com.sun.star.awt.SystemPointer import WAIT, ARROW, HAND

# Note only uncomment if need a msgbox for debugging
# If uncommented, you must open the python script organizer (alt + shift + f11) EVERY time you open the document
# from apso_utils import msgbox


########################################################################################################################
# Logger
# Structured logger.
########################################################################################################################
class LogLevel(Enum):
    CRITICAL = "CRITICAL"
    ERROR = "ERROR"
    WARN = "WARN"
    INFO = "INFO"
    DEBUG = "DEBUG"
    TRACE = "TRACE"


class Logger:
    def __init__(self, level=LogLevel.WARN, context = {}):
        self.rank = {
            LogLevel.CRITICAL: 100,
            LogLevel.ERROR: 200,
            LogLevel.WARN: 300,
            LogLevel.INFO: 400,
            LogLevel.DEBUG: 500,
            LogLevel.TRACE: 600,
        }

        self.level = level
        self.context = context

    def log(self, level, message):
        if not level in self.rank:
            raise ValueError(F"Unknown log level {level}")

        if self.rank[level] <= self.rank[self.level]:
            output = {
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "msg": message,
                "lvl": level.value,
                "ctx": self.context,
            }

            output_json = json.dumps(output)

            print(output_json)

    def critical(self, message):
        self.log(LogLevel.CRITICAL, message)

    def error(self, message):
        self.log(LogLevel.ERROR, message)

    def warn(self, message):
        self.log(LogLevel.WARN, message)

    def info(self, message):
        self.log(LogLevel.INFO, message)

    def debug(self, message):
        self.log(LogLevel.DEBUG, message)

    def trace(self, message):
        self.log(LogLevel.TRACE, message)

    def with_keys(self, **kwargs):
        return Logger(level = self.level, context = self.context | kwargs)

########################################################################################################################
# ControllerWrapper
# Wrapper for the controller and model
# Has some generic helper methods, not specific to Accounting/Journal Processing.
########################################################################################################################
class ControllerWrapper:
    def __init__(self, logger):
        self.logger = logger.with_keys(class_name = "ControllerWrapper")

        logger = self.logger.with_keys(method="__init__")
        logger.trace("ControllerWrapper __init__ start")

        self.model = XSCRIPTCONTEXT.getDocument()
        if self.model is None:
            logger.error("model is None")
            return

        self.controller = self.model.getCurrentController()
        if self.controller is None:
            logger.error("controller is None")
            return

        self.frame = self.controller.getFrame()
        if self.frame is None:
            logger.error("frame is None")
            return

        self.window = self.frame.getComponentWindow()
        if self.window is None:
            logger.error("window is None")
            return

        self.context = uno.getComponentContext()
        if self.context is None:
            logger.error("context is None")
            return

        self.pointer = self.context.ServiceManager.createInstanceWithContext("com.sun.star.awt.Pointer", self.context)
        if self.pointer is None:
            logger.error("pointer is None")
            return

        logger.trace("ControllerWrapper __init__ end")

    def get_active_sheet_name(self):
        logger = self.logger.with_keys(method="get_active_sheet_name")

        active_sheet = self.controller.getActiveSheet()
        if active_sheet is None:
            logger.warn("Active sheet is None")
            return

        sheet_name = active_sheet.getName()
        if sheet_name is None:
            logger.warn("sheet_name is none")

        return sheet_name

    def get_sheet_names(self):
        names = [name for name in self.model.Sheets.ElementNames]
        return names

    def ui_updates_stop(self):
        logger = self.logger.with_keys(method = "ui_updates_stop")
        logger.trace("method running")

        self.model.lockControllers()

    def ui_updates_start(self):
        logger = self.logger.with_keys(method = "ui_updates_start")
        logger.trace("method running")

        self.model.unlockControllers()

    def windowDisable(self):
        self.window.setEnable(False)

    def windowEnable(self):
        self.window.setEnable(True)

    def get_empty_cell_ranges(self):
        return self.model.createInstance("com.sun.star.sheet.SheetCellRanges")

    def set_pointer_type(self, pointer_type):
        self.pointer.setType(pointer_type)

########################################################################################################################
# JournalReader
# Reads the data from the Journal, creates Journal object with JournalEntries.
########################################################################################################################
class JournalReader:
    def __init__(self, logger, controller, journal_sheet):
        self.logger = logger.with_keys(class_name = "JournalReader")
        logger = self.logger.with_keys(method = "__init__")

        logger.trace("JournalReader __init__ start")

        self.controller = controller
        self.journal_sheet = journal_sheet

        self.COL_REC = 0
        self.COL_BUD = 1
        self.COL_CHECK = 2
        self.COL_DATE = 3
        self.COL_TRANS = 4
        self.COL_FROM = 5
        self.COL_TO = 6
        self.COL_AMOUNT = 7

        logger.trace("JournalReader __init__ end")

    def read(self):
        logger = self.logger.with_keys(method = "read")

        used_range = SheetHelper.get_used_range(logger, self.journal_sheet)
        logger.trace("Got Used Range")

        data_array = used_range.DataArray  # data_array is a nested tuple (rows, then columns)

        journal = Journal()

        for row_index, row in enumerate(data_array):
            if row_index != 0:  # Skip titles
                entry = self.get_entry(row_index, row)
                journal.add(entry)

        return journal

    def get_entry(self, index, row):
        logger = self.logger.with_keys(method = "get_entry")

        entry = JournalEntry(
            index,
            row[self.COL_REC],
            row[self.COL_BUD],
            row[self.COL_CHECK],
            row[self.COL_DATE],
            row[self.COL_TRANS],
            row[self.COL_FROM],
            row[self.COL_TO],
            row[self.COL_AMOUNT],
        )
        if (index - 1) % 1000 == 0:
            logger = logger.with_keys(entry = vars(entry))
            logger.trace("Created JournalEntry")

        return entry

    @staticmethod
    def get_journal_sheet(controller, logger, sheet_info):
        logger = logger.with_keys(class_name= "JournalReader", method = "get_journal_sheet")

        if sheet_info.group is None:
            logger.warn("Unable to parse sheet_name")
            return

        journal_name = f"{sheet_info.group}_Journal"
        logger = logger.with_keys(journal_name = journal_name)

        if journal_name == sheet_info.sheet_name:
            logger.warn("Active sheet is a journal sheet")
            return

        journal_sheet = SheetHelper.get_sheet_by_name(logger, controller, journal_name)

        return journal_sheet

class SheetInfo():
    def __init__(self, sheet_name):
        group, type = SheetHelper.parse_sheet_name(sheet_name)

        self.sheet_name = sheet_name
        self.group = group
        self.type = type

########################################################################################################################
# ErrorInfo
# Info about errors to help find them on the sheet
########################################################################################################################
class ErrorInfo:
    def __init__(self, row, date, message, from_acct, to_acct, amount):
        self.row = row
        self.date = date
        self.message = message
        self.from_acct = from_acct
        self.to_acct = to_acct
        self.amount = amount

########################################################################################################################
# Journal
# Holds Journal data
# Helper classes
########################################################################################################################
class JournalEntry:
    def __init__(self, index, rec, bud, check, date, trans, from_acct, to_acct, amount):
        self.index = index
        self.rec = rec
        self.bud = bud
        self.check = check
        self.date = date
        self.trans = trans
        self.from_acct = ensure_string(from_acct)
        self.to_acct = ensure_string(to_acct)
        self.amount_dollars = amount

class Journal:
    def __init__(self):
        self.entries = []

    def add(self, entry):
        self.entries.append(entry)

########################################################################################################################
# Account
########################################################################################################################
class AccountSettings:
    def __init__(self, name, sign, short, round, type, base):
        self.name = ensure_string(name)
        self.sign = int(sign)
        self.short = (short == 1)
        self.round = int(round)
        self.type = ensure_string(type)
        self.base = ensure_string(base)

class Accounts:
    def __init__(self, logger, controller):
        self.ACCOUNTS_SHEET_NAME = "Accounts"

        self.logger = logger.with_keys(
            class_name = "Accounts",
            method = "__init__",
        )
        logger = self.logger.with_keys(account_sheet_name = self.ACCOUNTS_SHEET_NAME)

        logger.trace("Accounts __init__ start")

        self.controller = controller

        self.COL_ACCOUNT = 0
        self.COL_SORT = 1
        self.COL_SIGN = 2
        self.COL_SHORT = 3
        self.COL_ROUND = 4
        self.COL_TYPE = 5
        self.COL_BASE = 6

        accounts_sheet = SheetHelper.get_sheet_by_name(logger, self.controller, self.ACCOUNTS_SHEET_NAME)

        used_range = SheetHelper.get_used_range(logger, accounts_sheet)
        logger.trace("Got Used Range")

        data_array = used_range.DataArray  # data_array is a nested tuple (rows, then columns)

        self.accounts = {}

        for row_index, row in enumerate(data_array):
            if row_index != 0:  # skip titles
                account = self.get_account(row)

                if account.name is None or account.name == "":
                    continue

                if account.name in self.accounts:
                    raise ValueError(f"Duplicate Account '{account.name}' at {row_index}")


                # logger.with_keys(account = vars(account)).trace("Got Account")
                self.accounts[account.name] = account

        logger.trace("Accounts __init__ end")

    def get_account(self, row):
        entry = AccountSettings(
            row[self.COL_ACCOUNT],
            row[self.COL_SIGN],
            row[self.COL_SHORT],
            row[self.COL_ROUND],
            row[self.COL_TYPE],
            row[self.COL_BASE],
        )

        return entry

########################################################################################################################
# Settings
########################################################################################################################
class Settings:
    def __init__(self, logger, controller):
        self.SHEET_NAME = "Settings"

        self.logger = logger.with_keys(
            class_name = "Settings",
            method = "__init__",
        )
        logger = self.logger.with_keys(account_sheet_name = self.SHEET_NAME)

        logger.trace("Settings __init__ start")

        self.controller = controller

        accounts_sheet = SheetHelper.get_sheet_by_name(logger, self.controller, self.SHEET_NAME)

        used_range = SheetHelper.get_used_range(logger, accounts_sheet)
        logger.trace("Got Used Range")

        data_array = used_range.DataArray  # data_array is a nested tuple (rows, then columns)

        self.settings = {}

        for row_index, row in enumerate(data_array):
            if row_index == 0:
                pass
            else:
                self.settings[row[0]] = row[1]

        self.BUDGET_MONTH = "BudgetMonth"
        self.PAY_DAY = "PayDay"

        logger.trace("Settings __init__ end")

    def getBudgetMonth(self):
        return self.getDateSetting(self.BUDGET_MONTH)

    def getPayDay(self):
        return self.getDateSetting(self.PAY_DAY)

    def getDateSetting(self, key):
        if key in self.settings:
            date_setting = self.settings[key]
            if isinstance(date_setting, numbers.Number):
                return convert_excel_date_to_date(date_setting)
            else:
                return None
        else:
            return None

########################################################################################################################
# Budget Lookup
########################################################################################################################
class BudgetLookup:
    def __init__(self, logger, settings):
        self.SHEET_NAME = "BudgetLookup"

        self.logger = logger.with_keys(
            class_name = "BudgetLookup",
            method = "__init__",
        )
        logger = self.logger.with_keys(account_sheet_name = self.SHEET_NAME)

        logger.trace("BudgetLookup __init__ start")

        self.settings = settings

        lookup, budget_month, extra_payday = self.calculate_values(logger)
        self.lookup = lookup
        self.budget_month = budget_month
        self.extra_payday = extra_payday

        logger.trace("BudgetLookup __init__ end")

    def calculate_values(self, logger):
        raw_budget_month = self.settings.getBudgetMonth()
        budget_month = self.set_day_of_month(raw_budget_month, 1)

        lookup = {}
        for month_index in range(-1,3):
            month = self.add_months(budget_month, month_index)

            first_day, last_day = self.get_month_first_and_last_date(logger, month)

            logger.with_keys(
                budget_month=budget_month.strftime("%Y%m%d %A"),
                month_index = month_index,
                month = month.strftime("%Y%m%d %A"),
                first_day=first_day.strftime("%Y%m%d %A"),
                last_day=last_day.strftime("%Y%m%d %A"),
            ).trace("Calculated first and last day of month")

            for day in range(1,29):
                key = day + ((month_index + 1) * 100)
                value = first_day + timedelta(days = day - 1)

                lookup[key] = value

            key = 30 + ((month_index + 1) * 100)
            value = last_day

            lookup[key] = value

        first_friday = self.find_next_day_of_week(budget_month, 4) # 4 = friday

        first_payday = first_friday
        if not self.is_payday(first_friday):
            first_payday = first_friday + timedelta(days = 7)

        third_payday = first_payday + timedelta(days = 28)

        _, last_day = self.get_month_first_and_last_date(logger, budget_month)

        extra_payday = third_payday <= last_day
        logger.with_keys(
            extra_payday = extra_payday,
            third_payday = third_payday.strftime("%Y%m%d %A"),
            last_day = last_day.strftime("%Y%m%d %A"),
        ).trace("extra_payday calculation values")

        start_of_week_one = first_payday + timedelta(days = -5)

        lookup_date = start_of_week_one
        for week in range(1,7):
            for day in range(1,8):
                lookup[f"W{week}.{day}"] = lookup_date
                lookup_date = lookup_date + timedelta(days = 1)

        return lookup, budget_month, extra_payday

    def is_payday(self, date_value):
    # def is_payday(target_date, anchor_payday):
        anchor_payday = self.settings.getPayDay()
        delta = (date_value - anchor_payday).days

        # Biweekly pay occurs every 14 days
        return delta % 14 == 0

    def find_next_day_of_week(self, date_value, day_of_week):
        search_date = date_value
        while search_date.weekday() != day_of_week:
            search_date = search_date + timedelta(days = 1)

        return search_date

    def add_months(self, date_value, month_count):
        # note this only works reliably for dates with day 1-28
        # good enough for our purposes since we almost always use the 1st
        new_month = ((date_value.month + month_count - 1) % 12) + 1
        return date(date_value.year, new_month, date_value.day)

    def set_day_of_month(self, date_value, day_of_month):
        return date(date_value.year, date_value.month, day_of_month)

    def get_month_first_and_last_date(self, logger, month):
        logger = logger.with_keys(method = "get_month_first_and_last_date")

        _, days_in_month = calendar.monthrange(month.year, month.month)
        first_day = self.set_day_of_month(month, 1)
        last_day = self.set_day_of_month(month, days_in_month)

        logger.with_keys(
            month=month.isoformat(),
            first_day=first_day.isoformat(),
            last_day=last_day.isoformat()
        )
        logger.trace("Calculated first and last day of month")

        return first_day, last_day

########################################################################################################################
# Report Generator
########################################################################################################################
class ReportGeneratorBase():
    def __init__(self, logger, journal, accounts, settings):
        self.logger = logger.with_keys(class_name = ReportGeneratorBase)
        logger = logger.with_keys(method = "__init__")

        logger.trace("ReportGeneratorBase __init__ start")

        self.journal = journal
        self.accounts = accounts
        self.settings = settings
        self.report = None
        self.custom_accounts = {}

        logger.trace("ReportGeneratorBase __init__ end")

    def generate(self):
        logger = self.logger.with_keys(method = "generate")

        logger.trace("Generating Report Journal Entries")
        self.generate_report_entries()

        logger.trace("Generating Report Account Balances")
        self.generate_account_balances()

        logger.trace("Generating Custom Data")
        self.generate_custom_data()

        logger.trace

    def generate_custom_data(self):
        pass

    def generate_report_entries(self):
        logger = self.logger.with_keys(method = "generate_report_entries")

        self.report = Report(logger)

        if len(self.journal.entries) == 0:
            logger.warn("No journal entries")

        for journal_index, journal_entry in enumerate(self.journal.entries):
            report_entry = JournalEntry(
                journal_entry.index,
                journal_entry.rec,
                journal_entry.bud,
                journal_entry.check,
                self.date_translate(journal_entry.date),
                self.trans_translate(journal_entry.trans),
                self.account_translate(journal_entry, journal_entry.from_acct, True),
                self.account_translate(journal_entry, journal_entry.to_acct, False),
                self.amount_dollars_translate(journal_entry.amount_dollars, journal_entry),
            )

            self.report.add(report_entry)

            if (journal_index - 1) % 1000 == 0:
                logger.with_keys(report_entry = vars(report_entry)).trace("Created Report Entry")
                logger.with_keys(journal_entry = vars(journal_entry)).trace("Original Journal Entry")

    def validate_account(self, entry, name, is_debit):
        name_required = isinstance(entry.amount_dollars, numbers.Number)
        if name_required and (name is None or name == ""):
            name = "<blank>"
        
        if name is not None and name != "" and name not in self.accounts.accounts:
            if is_debit:
                from_acct = name
                to_acct = ""
                error = "Bad FROM"
            else:
                from_acct = ""
                to_acct = name
                error = "Bad TO"

            return ErrorInfo(
                entry.index + 1, # row is index plus one
                entry.date,
                f"{error} - {entry.trans}",
                from_acct,
                to_acct,
                entry.amount_dollars
            )

    def date_translate(self, date):
        return date

    def trans_translate(self, trans):
        return trans

    def account_translate(self, entry, account, is_debit):
        self.report.add_error(
            self.validate_account(entry, account, is_debit)
        )

        return account

    def amount_dollars_translate(self, amount_dollars, journal_entry):
        return amount_dollars

    def generate_account_balances(self):
        logger = self.logger.with_keys(method = "generate_account_balances")

        balances = {}

        for report_entry in self.report.entries:
            self.create_and_add_balance_entry(balances, report_entry, report_entry.from_acct, self.calculate_trans_amt)

            if report_entry.from_acct != report_entry.to_acct:
                self.create_and_add_balance_entry(balances, report_entry, report_entry.to_acct, self.calculate_trans_amt)

            for acct_name, trans_amt_calculator in self.custom_accounts:
                custom_entry = self.create_account_balance_entry(report_entry, acct_name, trans_amt_calculator)
                self.add_balance_entry(balances, custom_entry)

        logger.trace("Calculating Extrema")
        self.calculate_extrema(balances)

        logging_acct = "FED.2023"
        # logging_acct = "4753"
        # logging_acct = "401K.CLGX"
        if logging_acct in balances:
            for entry in balances[logging_acct]:
                logger.with_keys(entry = vars(entry)).trace("Account Balance Entry")
        else:
            logger.with_keys(logging_acct = logging_acct).trace("Logging account not in balances")

        self.report.balances = balances

    def create_and_add_balance_entry(self, balances, report_entry, acct, calculate_trans_amt):
        if acct in self.accounts.accounts:
            balance_entry = self.create_account_balance_entry(report_entry, acct, calculate_trans_amt)
            self.add_balance_entry(balances, balance_entry)

    def calculate_extrema(self, balances):
        for account, entries in balances.items():
            entry_count = len(entries)
            for entry_index, entry in enumerate(entries):
                if entry_index == 0:
                    prior_balance = 0
                else:
                    prior_balance = entries[entry_index - 1].balance_cents

                curr_balance = entry.balance_cents

                if entry_index == entry_count - 1:
                    next_balance = prior_balance
                else:
                    next_balance = entries[entry_index + 1].balance_cents

                if prior_balance < curr_balance and next_balance < curr_balance:
                    entry.extrema = ExremaType.MAXIMA
                elif prior_balance > curr_balance and next_balance > curr_balance:
                    entry.extrema = ExremaType.MINIMA
                else:
                    entry.extrema = ExremaType.NONE

    def add_balance_entry(self, balances, entry):
        if entry is not None:
            if entry.acct_name not in balances:
                balances[entry.acct_name] = []

            acct_entries = balances[entry.acct_name]
            if len(acct_entries) == 0:
                entry.balance_cents = entry.trans_amt_cents
            else:
                prior_entry = acct_entries[-1]
                entry.balance_cents = prior_entry.balance_cents + entry.trans_amt_cents

            acct_entries.append(entry)

    def create_account_balance_entry(self, report_entry, acct_name, trans_amt_calculator):
        trans_amt = trans_amt_calculator(report_entry, acct_name)
        if trans_amt == 0:
            return None

        entry = AccountBalanceEntry(
            acct_name,
            report_entry.index,
            trans_amt,
        )

        return entry

    def calculate_trans_amt(self, report_entry, acct_name):
        logger = self.logger.with_keys(method = "calculate_trans_amt")

        # isinstance(value, numbers.Number):
        if not isinstance(report_entry.amount_dollars, numbers.Number):
            # logger.with_keys(
            #     amount_dollars = report_entry.amount_dollars,
            #     report_entry = vars(report_entry)
            # ).trace("Amount Dollars is not a number")

            return 0

        trans_amt = 0
        sign = self.accounts.accounts[acct_name].sign
        if report_entry.from_acct == acct_name:
            trans_amt -= convert_dollars_to_cents(report_entry.amount_dollars) * sign
        if report_entry.to_acct == acct_name:
            trans_amt += convert_dollars_to_cents(report_entry.amount_dollars) * sign

        return trans_amt

class AccountBalanceEntry:
    def __init__(self, acct_name, index, trans_amt_cents):
        self.acct_name = acct_name
        self.index = index
        self.trans_amt_cents = trans_amt_cents
        self.balance_cents = None
        self.extrema = ExremaType.UNKNOWN


class ExremaType(str, Enum):
    UNKNOWN = "UNKNOWN"
    MAXIMA = "MAXIMA"
    NONE = "NONE"
    MINIMA = "MINIMA"


class BalReportGenerator(ReportGeneratorBase):
    def __init__(self, logger, journal, accounts, settings):
        super().__init__(logger, journal, accounts, settings)

        self.logger = self.logger.with_keys(class_name = "BalReportGenerator")

        logger = logger.with_keys(method="__init__")

        logger.trace("BalReportGenerator __init__ start")

        logger.trace("BalReportGenerator __init__ end")

class CopyReportGenerator(ReportGeneratorBase):
    def __init__(self, logger, journal, accounts, settings):
        super().__init__(logger, journal, accounts, settings)

        self.logger = self.logger.with_keys(class_name = "CopyReportGenerator")

        logger = logger.with_keys(method="__init__")

        logger.trace("CopyReportGenerator __init__ start")

        self.lookup = BudgetLookup(self.logger, self.settings)
        self.month_string = self.lookup.budget_month.strftime("%b")
        self.excel_lookup = self.create_excel_lookup()

        logger.trace("CopyReportGenerator __init__ end")

    def create_excel_lookup(self):
        excel_lookup = {}

        for key, value in self.lookup.lookup.items():
            excel_lookup[key] = convert_date_to_excel_date(value)

        return excel_lookup

    def date_translate(self, date):
        super_date = super().date_translate(date)

        if super_date in self.excel_lookup:
            return self.excel_lookup[super_date]
        else:
            return super_date

    def trans_translate(self, trans):
        super_trans = super().trans_translate(trans)

        if super_trans is not None and super_trans != "":
            return f"{super_trans} ({self.month_string})"
        else:
            return super_trans

    def generate_custom_data(self):
        super().generate_custom_data()

        self.report.data["lookup"] = self.lookup

class RecReportGenerator(ReportGeneratorBase):
    def __init__(self, logger, journal, accounts, settings):
        super().__init__(logger, journal, accounts, settings)

        self.logger = self.logger.with_keys(class_name = "RecReportGenerator")

        logger = logger.with_keys(method="__init__")

        logger.trace("RecReportGenerator __init__ start")


        logger.trace("RecReportGenerator __init__ end")

    def account_translate(self, entry, account, is_debit):
        logger = self.logger.with_keys(method = "account_translate")

        super_account = super().account_translate(entry, account, is_debit)

        logger = logger.with_keys(account = account, super_account = super_account).trace("got super account")

        if super_account not in self.accounts.accounts:
            return super_account

        base_account = self.accounts.accounts[super_account].base

        self.report.add_error(
            self.validate_account(entry, base_account, is_debit)
        )

        return base_account

    def amount_dollars_translate(self, amount_dollars, journal_entry):
        super_amount_dollars = super().amount_dollars_translate(amount_dollars, journal_entry)

        if journal_entry.rec != 1 and isinstance(amount_dollars, numbers.Number):
            return 0.0
        else:
            return amount_dollars

class Report:
    def __init__(self, logger):
        self.logger = logger.with_keys(class_name = "Report")

        logger = logger.with_keys(method="__init__")

        logger.trace("Report __init__ start")

        self.entries = []
        self.balances = {}
        self.data = {}
        self.errors = []

        logger.trace("Report __init__ end")

    def add(self, entry):
        self.entries.append(entry)

    # error is an ErrorInfo
    def add_error(self, error):
        if error is not None:
            self.errors.append(error)

########################################################################################################################
# Report Updater
########################################################################################################################
class ReportUpdaterBase():
    def __init__(self, logger, controller, report_sheet, report_generator):
        self.logger = logger.with_keys(class_name = ReportUpdaterBase)
        logger = logger.with_keys(method = "__init__")

        logger.trace("ReportUpdaterBase __init__ start")

        self.controller = controller
        self.sheet = report_sheet
        
        self.accounts = report_generator.accounts.accounts
        self.balances = report_generator.report.balances
        self.entries = report_generator.report.entries
        self.errors = report_generator.report.errors
        self.data = report_generator.report.data

        self._ENTRY_COL_REC = 0
        self._ENTRY_COL_BUD = 1
        self._ENTRY_COL_CHECK = 2
        self._ENTRY_COL_DATE = 3
        self._ENTRY_COL_TRANS = 4
        self._ENTRY_COL_FROM_ACCT = 5
        self._ENTRY_COL_TO_ACCT = 6
        self._ENTRY_COL_AMOUNT = 7

        self._FILTER_COL_ACCOUNT = 8
        self._FILTER_COL_DEBIT = 9
        self._FILTER_COL_CREDIT = 10

        self._BALANCES_COL_FIRST = 11

        logger.trace("ReportUpdaterBase __init__ end")

    def update(self):
        logger = self.logger.with_keys(method = "update")

        used_range_address = self.get_used_range_address()

        self.clear_current_report(used_range_address)
        used_range_address = self.get_used_range_address()

        data_array = self.build_data_array(used_range_address)

        self.set_report_data(data_array, used_range_address)
        used_range_address = self.get_used_range_address()

        self.apply_formatting(used_range_address)

        self.validate()
        self.show_errors(used_range_address)

    def add_error(self, error):
        if error is not None:
            self.errors.append(error)

    def show_errors(self, used_range_address):
        logger = self.logger.with_keys(method = "show_errors")
        
        # Indicator Test
        indicator_text = "Trans"
        message_count = len(self.errors)

        if message_count != 0:
            indicator_text = f"Trans - Messages ({message_count})"

        data_array = DataArrayHelper.create_empty_data_array(1, 1)
        data_array[0][0] = indicator_text

        data_range = self.sheet.getCellRangeByPosition(self._ENTRY_COL_TRANS, 0, self._ENTRY_COL_TRANS, 0)
        data_range.setDataArray(data_array)

        # Error Messages
        if message_count == 0:
            return

        data_array = DataArrayHelper.create_empty_data_array(6, message_count)

        for error_index, error in enumerate(self.errors):
            data_array[error_index][0] = error.row
            data_array[error_index][1] = error.date
            data_array[error_index][2] = error.message
            data_array[error_index][3] = error.from_acct
            data_array[error_index][4] = error.to_acct
            data_array[error_index][5] = error.amount

        start_col = self._ENTRY_COL_CHECK
        start_row = used_range_address.EndRow + 2
        end_col = start_col + 5
        end_row = start_row + message_count - 1

        logger.with_keys(
            start_col=start_col,
            start_row=start_row,
            end_col=end_col,
            end_row=end_row
        ).trace("update coordinates")

        # update worksheet
        data_range = self.sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
        data_range.setDataArray(data_array)

        # highlight errors:
        color = ReportColors.Error

        self.set_color_by_position(color, start_col, start_row, end_col, end_row)

    def get_used_range_address(self):
        logger = self.logger.with_keys(method = "get_used_range")

        used_range_address = SheetHelper.get_used_range_address(logger, self.sheet)

        logger.with_keys(
            start_col = used_range_address.StartColumn,
            start_row = used_range_address.StartRow,
            end_col = used_range_address.EndColumn,
            end_row = used_range_address.EndRow
        ).trace("Used Range in Report Before Update")

        return used_range_address

    def clear_current_report(self, used_range_address):
        logger = self.logger.with_keys(method = "clear_current_report")

        rows_to_delete = used_range_address.EndRow - used_range_address.StartRow
        if rows_to_delete == 0:
            logger.with_keys(rows_to_delete=rows_to_delete).trace("Nothing to delete, skipping")
            return

        logger.with_keys(rows_to_delete = rows_to_delete).trace("Before deleting rows in report")
        SheetHelper.delete_rows_range(self.sheet, 1, rows_to_delete)

    def build_data_array(self, used_range_address):
        data_array_list = self.create_empty_data_array(used_range_address)
        self.build_data_array_entries(data_array_list)

        if used_range_address.EndColumn >= self._FILTER_COL_CREDIT:
            self.build_data_array_filter(data_array_list)

        if used_range_address.EndColumn >= self._BALANCES_COL_FIRST:
            self.build_data_array_balances(data_array_list, used_range_address)

        return self.convert_data_array_to_tuples(data_array_list)

    def create_empty_data_array(self, used_range_address):
        width = used_range_address.EndColumn + 1
        height = len(self.entries)

        return DataArrayHelper.create_empty_data_array(width, height)

    def build_data_array_entries(self, data_array_list):
        report_entries = self.entries

        for index, entry in enumerate(report_entries):
            data_array_list[index][self._ENTRY_COL_REC] = self.ensure_flag_is_int(entry.rec)
            data_array_list[index][self._ENTRY_COL_BUD] = self.ensure_flag_is_int(entry.bud)
            data_array_list[index][self._ENTRY_COL_CHECK] = entry.check
            data_array_list[index][self._ENTRY_COL_DATE] = entry.date
            data_array_list[index][self._ENTRY_COL_TRANS] = entry.trans
            data_array_list[index][self._ENTRY_COL_FROM_ACCT] = entry.from_acct
            data_array_list[index][self._ENTRY_COL_TO_ACCT] = entry.to_acct
            data_array_list[index][self._ENTRY_COL_AMOUNT] = entry.amount_dollars

    def ensure_flag_is_int(self, flag):
        if isinstance(flag, numbers.Number):
            return int(flag)
        else:
            return flag

    def build_data_array_filter(self, data_array_list):
        logger = self.logger.with_keys(method = "build_data_array_filter")

        account_cell = self.sheet.getCellByPosition(self._FILTER_COL_ACCOUNT, 0)
        account = account_cell.String
        logger.with_keys(account = account).trace("Got filter account")

        report_entries = self.entries

        for index, entry in enumerate(report_entries):
            if entry.from_acct == account:
                data_array_list[index][self._FILTER_COL_ACCOUNT] = 1
                data_array_list[index][self._FILTER_COL_DEBIT] = entry.amount_dollars

            if entry.to_acct == account:
                data_array_list[index][self._FILTER_COL_ACCOUNT] = 1
                data_array_list[index][self._FILTER_COL_CREDIT] = entry.amount_dollars

    def build_data_array_balances(self, data_array_list, used_range_address):
        logger = self.logger.with_keys(method = "build_data_array_balances")

        balances = self.balances
        account_names_array = self.get_account_names_array(used_range_address)

        row_count = len(self.entries)
        logger.with_keys(row_count = row_count).trace('got row count (number of report entries)')

        logger.with_keys(accounts_count = len(account_names_array[0])).trace("count of accounts")

        for account_index, account_data in enumerate(account_names_array[0]):
            account = ensure_string(account_data)
            logger = logger.with_keys(account_index = account_index, account = account)
            # logger.trace("calculating balances for account")

            if account in balances:
                # logger.trace("filling out accounts")
                balance_data = balances[account]
                first_balance = balance_data[0]

                # Fill in leading zero's
                if first_balance.index > 1:
                    # logger.trace("balances - fill in initial zeros")
                    for row_index in range(0, first_balance.index - 1):
                        data_array_list[row_index][account_index + self._BALANCES_COL_FIRST] = 0.0

                last_balance_index = len(balance_data) - 1
                for balance_index, balance in enumerate(balance_data):
                    logger = logger.with_keys(balance_index = balance_index)
                    # logger.trace("processing nth balance for account")

                    if balance_index == last_balance_index:
                        end_row = row_count
                    else:
                        end_row = balance_data[balance_index + 1].index - 1

                    for row_index in range(balance.index - 1, end_row):
                        data_array_list[row_index][account_index + self._BALANCES_COL_FIRST] = balance.balance_cents / 100

    def convert_data_array_to_tuples(self, data_array_list):
        logger = self.logger.with_keys(method = "convert_data_array_to_tuples")

        entry_tuples = []
        for entry in data_array_list:
            entry_tuples.append(tuple(entry))

        data_array = tuple(entry_tuples)

        logger.with_keys(width=len(data_array[0]),length=len(data_array)).trace("data_array size")

        # Dumper.dump("data array tuples", data_array)

        return data_array

    def set_report_data(self, data_array, used_range_address):
        data_range = self.sheet.getCellRangeByPosition(0, 1, used_range_address.EndColumn, len(data_array))
        data_range.setDataArray(data_array)

    def apply_formatting(self, used_range_address):
        logger = self.logger.with_keys(method = "apply_formatting")

        balances = self.balances

        if used_range_address.EndColumn < self._BALANCES_COL_FIRST:
            return

        account_names_array = self.get_account_names_array(used_range_address)

        row_count = len(self.entries)
        logger.with_keys(row_count = row_count).trace('got row count (number of report entries)')

        logger.with_keys(accounts_count = len(account_names_array[0])).trace("count of accounts")

        account_settings = self.accounts
        global_short_flag = False
        short_indexes = set()

        short_color = ReportColors.Error

        for account_index, account_data in enumerate(account_names_array[0]):
            account = ensure_string(account_data)
            logger = logger.with_keys(account_index = account_index, account = account)
            # logger.trace("applying formatting for account")

            if account in balances:
                # logger.trace("account found in balances")
                balance_data = balances[account]

                curr_account_settings = account_settings[account]
                can_be_short = curr_account_settings.short

                last_balance_index = len(balance_data) - 1
                for balance_index, balance in enumerate(balance_data):
                    logger = logger.with_keys(balance_index = balance_index)
                    # logger.trace("processing nth balance for account")

                    is_short = can_be_short and balance.balance_cents < 0

                    if is_short:
                        global_short_flag = True
                        short_indexes.add(balance.index)

                        account_col = account_index + self._BALANCES_COL_FIRST

                        start_row = balance.index + 1
                        if balance_index == last_balance_index:
                            end_row = row_count
                        else:
                            end_row = balance_data[balance_index + 1].index - 1

                        if end_row >= start_row:
                            self.set_color_by_position(
                                short_color,
                                account_col,
                                start_row,
                                account_col,
                                end_row,
                            )

                    extrema_color = ReportColors.MinMax[is_short][balance.extrema]
                    if extrema_color is not None:
                        self.set_color_by_position(extrema_color, account_index + self._BALANCES_COL_FIRST, balance.index)

        if global_short_flag:
            short_indexes.add(0) # tag the Amount column header.
        else:
            self.set_color_by_position(-1, self._ENTRY_COL_AMOUNT, 0)


        for short_index in short_indexes:
            self.set_color_by_position(short_color, self._ENTRY_COL_AMOUNT, short_index)

        report_entries = self.entries
        reconciled_entries = [entry for entry in report_entries if entry.rec == 1]
        unreconciled_entries = [entry for entry in report_entries if entry.rec != 1]

        if len(reconciled_entries) != 0:
            last_reconciled_index = reconciled_entries[-1].index
            first_unreconciled_index = unreconciled_entries[0].index

            logger.with_keys(last_reconciled_index = last_reconciled_index, first_unreconciled_index = first_unreconciled_index).trace("got reconciliation rows")

            # mark starting reconciled
            reconciled_color = ReportColors.HighHighlight
            self.set_color_by_position(
                reconciled_color,
                self._ENTRY_COL_REC,
                1,
                self._ENTRY_COL_REC,
                first_unreconciled_index - 1,
            )

            entries_to_check = [
                entry for entry in report_entries
                if entry.index >= first_unreconciled_index and entry.index < last_reconciled_index
            ]

            for entry in entries_to_check:
                if entry.rec != 1:
                    unreconciled_color = ReportColors.LowHighlight
                    self.set_color_by_position(
                        unreconciled_color,
                        self._ENTRY_COL_REC,
                        entry.index,
                    )


    def validate(self):
        pass

    def set_color_by_position(self, color, start_col, start_row, end_col = None, end_row = None, toggle = False):
        SheetHelper.set_range_background_color(self.sheet, color, start_col, start_row, end_col, end_row, toggle)

    def get_account_names_array(self, used_range_address):
        account_names_range = self.sheet.getCellRangeByPosition(
            self._BALANCES_COL_FIRST,
            0,
            used_range_address.EndColumn,
            0
        )

        return account_names_range.DataArray

    def toggle_processing_indicator(self):
        indicator_color = ReportColors.NeutralHighlight
        self.set_color_by_position(indicator_color, self._ENTRY_COL_TRANS, 0, toggle = True)


class BalReportUpdater(ReportUpdaterBase):
    def __init__(self, logger, controller, report_sheet, report_generator):
        super().__init__(logger, controller, report_sheet, report_generator)

        self.logger = self.logger.with_keys(class_name = "BalReportUpdater")
        logger = logger.with_keys(method="__init__")

        logger.trace("BalReportUpdater __init__ start")


        logger.trace("BalReportUpdater __init__ end")

    def validate(self):
        super().validate()

        logger = self.logger.with_keys(method = "validate")
        logger.trace("method_start in BalReportUpdater")

        account_settings = self.accounts
        balances = self.balances

        # find bad balances
        for account in account_settings.values():
            if account.name in balances:
                account_balance = balances[account.name]
                last_balance = account_balance[-1]

                # account.round is number of digits to round dollars, needs adjustment for cents
                round_digits_cents = 2 - account.round
                rounding_factor = pow(10, round_digits_cents)
                # rounding factor allows some accounts to round to cents, some to dollars
                # and some to 10,000,000 dollars (effectively ignore that account)
                rounded_balance = int(last_balance.balance_cents/rounding_factor)

                if rounded_balance != 0:
                    self.add_error(
                        ErrorInfo(
                            "",
                            "",
                            "Bad Final Balance",
                            account.name,
                            "",
                            last_balance.balance_cents / 100
                        )
                    )

class CopyReportUpdater(ReportUpdaterBase):
    def __init__(self, logger, controller, report_sheet, report_generator):
        super().__init__(logger, controller, report_sheet, report_generator)

        self.logger = self.logger.with_keys(class_name = "CopyReportUpdater")
        logger = logger.with_keys(method="__init__")

        logger.trace("CopyReportUpdater __init__ start")


        logger.trace("CopyReportUpdater __init__ end")

    def validate(self):
        super().validate()

        logger = self.logger.with_keys(method = "validate")
        logger.trace("method_start in BalReportUpdater")

        lookup_key = "lookup"
        if lookup_key not in self.data:
            self.add_error(
                ErrorInfo(
                    "",
                    "",
                    "lookup not found in data",
                    "",
                    "",
                    ""
                )
            )
            logger.error("lookup not found in data")
            return

        lookup = self.data[lookup_key]

        logger.with_keys(extra_payday = lookup.extra_payday).trace("got extra_payday from lookup")

        if not lookup.extra_payday:
            logger.trace("No Extra Payday")
            return

        logger.trace("Extra Payday")

        self.add_error(
            ErrorInfo(
                "",
                "",
                "This month has 3 paychecks",
                "",
                "",
                ""
            )
        )

class RecReportUpdater(ReportUpdaterBase):
    def __init__(self, logger, controller, report_sheet, report_generator):
        super().__init__(logger, controller, report_sheet, report_generator)

        self.logger = self.logger.with_keys(class_name = "RecReportUpdater")
        logger = logger.with_keys(method="__init__")

        logger.trace("RecReportUpdater __init__ start")


        logger.trace("RecReportUpdater __init__ end")


class ReportColors:
    # [Short][ExtrmaType]
    MinMax = {
        True: {
            ExremaType.MINIMA: int("C00000", 16),   # Burgundy
            ExremaType.NONE: int("FF0000", 16),     # Red
            ExremaType.MAXIMA: int("FF00FF", 16),   # Pink
        },
        False: {
            ExremaType.MINIMA: int("FFFF00", 16),   # Yellow
            ExremaType.NONE: None,                  # No Fill
            ExremaType.MAXIMA: int("00B050", 16),   # Med Green
        }
    }
    NeutralHighlight = int("00B0F0", 16)            # Light Blue
    LowHighlight = int("FFFF00", 16)                # Yellow
    HighHighlight = int("00B050", 16)               # Med Green
    Error = int("FF0000", 16)                       # Red

########################################################################################################################
# General helper functions
########################################################################################################################
def convert_dollars_to_cents(dollars):
    # Convert input to Decimal, then multiply by 100 and cast to int
    cents = int(Decimal(str(dollars)) * 100)
    return cents

def convert_excel_date_to_date(excel_date):
    python_date = datetime(1899, 12, 30) + timedelta(days=excel_date)
    return python_date.date()

def convert_date_to_excel_date(date_value):
    delta = date_value - datetime(1899, 12, 30).date()
    return delta.days

class DataArrayHelper:
    @staticmethod
    def create_empty_data_array(width, height, value = ""):
        data_array = []
        for row in range(0, height):
            data_array.append([value] * width)

        return data_array

def default_serialize(obj):
    if hasattr(obj, '__dict__'):
        return vars(obj)
        # Handle LibreOffice UNO objects or others by string conversion
    return str(obj)

class Dumper:
    @staticmethod
    def dump(filename, data):
        now = datetime.now()
        readable_filename_time = now.strftime("%Y_%m_%d__%H_%M_%S") + f"_{now.microsecond // 1000:03d}"

        with open(f"c:\\tmp\\{readable_filename_time}-{filename}.json",'w') as json_file:
            json.dump(
                data,
                json_file,
                indent = 4,
                default = default_serialize,
            )

def ensure_string(value):
    if isinstance(value, bool):
        return str(value)
    elif isinstance(value, numbers.Number):
        return str(int(value))
    else:
        return str(value)


class SheetHelper:
    @staticmethod
    def delete_rows_range(sheet, start_row, count):
        sheet.Rows.removeByIndex(start_row, count)

    @staticmethod
    def parse_sheet_name(name):
        parts = name.split('_')
        length = len(parts)
        if length != 2:
            return None, None
        else:
            return parts[0], parts[1]

    @staticmethod
    def get_used_range_address(logger, sheet):
        logger = logger.with_keys(class_name = "SheetHelper", method = "get_used_range_address")

        cursor = sheet.createCursor()
        # parameter: Expand (True) or select just the last cell (False)
        cursor.gotoEndOfUsedArea(True)

        return cursor.RangeAddress

    @staticmethod
    def get_used_range(logger, sheet):
        logger = logger.with_keys(class_name = "SheetHelper", method = "get_used_range")

        range_address = SheetHelper.get_used_range_address(logger, sheet)

        logger = logger.with_keys(
            start_col = range_address.StartColumn,
            start_row = range_address.StartRow,
            end_col = range_address.EndColumn,
            end_row = range_address.EndRow
        )
        logger.trace("Got range_address")

        used_range = sheet.getCellRangeByPosition(
            range_address.StartColumn,
            range_address.StartRow,
            range_address.EndColumn,
            range_address.EndRow)

        return used_range

    @staticmethod
    def set_range_background_color(sheet, color, start_col, start_row, end_col = None, end_row = None, toggle = False):
        if end_col == None:
            end_col = start_col
        if end_row == None:
            end_row = start_row

        range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)

        if toggle and range.CellBackColor != -1:
            color = -1

        range.CellBackColor = color

    @staticmethod
    def get_sheet_by_name(logger, controller, sheet_name):
        logger = logger.with_keys(class_name = "SheetHelper", method = "get_sheet_by_name", sheet_name = sheet_name)

        sheet_names = controller.get_sheet_names()
        logger = logger.with_keys(sheet_names = sheet_names)

        if sheet_name not in sheet_names:
            logger.warn("Sheet Not Found by Name")
            return

        named_sheet = controller.model.Sheets.getByName(sheet_name)

        return named_sheet


########################################################################################################################
# Macros for export/use
########################################################################################################################

def update_report(*args):
    start_time = time.perf_counter()

    log_level = LogLevel.WARN
    log_level = LogLevel.TRACE
    logger = Logger(level = log_level).with_keys(method = "update_report")

    try:

        logger.info("Setup")

        logger.trace("Create controller")
        controller = ControllerWrapper(logger)

        settings = Settings(logger, controller)

        logger.trace("Get journal Sheet")
        sheet_name = controller.get_active_sheet_name()
        if sheet_name is None:
            logger.error("sheet_name is None")
            return

        sheet_info = SheetInfo(sheet_name)
        logger.with_keys(sheet_info = vars(sheet_info)).trace("Got sheet_info")

        journal_sheet = JournalReader.get_journal_sheet(controller, logger, sheet_info)
        if journal_sheet is None:
            logger.warn("Journal Not Found, skipping")
            return

        logger.trace("Create journal Reader")
        journal_reader = JournalReader(logger, controller, journal_sheet)

        logger.trace("Read Journal")
        journal = journal_reader.read()

        logger.trace("Get Account Settings")
        accounts = Accounts(logger, controller)

        # determine report type
        report_generators = {
            "Bal":BalReportGenerator,
            "Copy":CopyReportGenerator,
            "Rec":RecReportGenerator,
        }

        report_generator = None
        if  sheet_info.type in report_generators:
            report_generator = report_generators[sheet_info.type](logger, journal, accounts, settings)
        else:
            logger.with_keys(
                sheet_info = vars(sheet_info),
                report_generators_keys = report_generators.keys()
            ).warn("Missing Report Generator for this sheet type")
            return

        report_generator.generate()

        # Dumper.dump("Report Journal", report_generator.report.entries)

        report_sheet = SheetHelper.get_sheet_by_name(logger, controller, sheet_name)

        report_updaters = {
            "Bal":BalReportUpdater,
            "Copy":CopyReportUpdater,
            "Rec":RecReportUpdater,
        }

        report_updater = None
        if  sheet_info.type in report_updaters:
            report_updater = report_updaters[sheet_info.type](logger, controller, report_sheet, report_generator)
        else:
            logger.with_keys(
                sheet_info = vars(sheet_info),
                report_generators_keys = report_generators.keys()
            ).warn("Missing Report Generator for this sheet type")
            return

        try:
            controller.ui_updates_stop()

            report_updater.update()
        finally:
            controller.ui_updates_start()
            report_updater.toggle_processing_indicator()

    finally:

        end_time = time.perf_counter()
        elapsed_time = end_time - start_time

        logger.with_keys(elapsed_time = f"{elapsed_time:.4f}").trace("FINISHED")

        # msgbox(f"Journal Sheet '{journal_sheet.Name}'",title="Info")


# Only the specified function will show in the Tools > Macro > Organize Macro dialog:
g_exportedScripts = (update_report,)
