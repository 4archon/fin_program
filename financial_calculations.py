import numpy_financial as np_f
import math


class FinIndicators:
    def __init__(self, revenue: tuple, capex: tuple, opex: tuple, investments_percentage: tuple,
                 number_of_credit_periods: tuple, credit_percentage: tuple, change_working_cap: tuple, useful_life: int,
                 discount_rate_fcff: float, discount_rate_fcfe: float):
        self.__revenue = revenue
        self.__capex = capex
        self.__opex = opex
        self.__costs = []
        self.__investments_percentage = investments_percentage
        self.__count_years = len(self.__revenue)
        self.__number_of_credit_periods = number_of_credit_periods
        self.__credit_percentage = credit_percentage
        self.__change_working_cap = change_working_cap
        self.__useful_life = useful_life
        self.__discount_rate_fcff = discount_rate_fcff
        self.__discount_rate_fcfe = discount_rate_fcfe
        for i in range(self.__count_years):
            self.__costs += [self.__capex[i] + self.__opex[i]]
        self.__costs = tuple(self.__costs)
        # кредитная часть
        self.__credits_matrix_payments = []
        self.__credits_matrix_summa = []
        self.__credits_matrix_overprice = []
        self.__credits_matrix_percentage_per_year = []
        self.__credits_matrix_balances = []
        # данные для вывода
        self.__reinvestment = [0]
        self.__equity_investments = [self.__costs[0] * self.__investments_percentage[0]]
        self.__credit_funds = [self.__costs[0] - self.__reinvestment[0] - self.__equity_investments[0]]
        self.__credit_balance = []
        # системные вызовы
        self.__credit_settlements()
        self.__conversion_tuple()

    def __credit_settlements(self):
        def percent_balance(credit_funds: (int, float), credit_percentage: float, number_of_credit_periods: int, i1):
            rate_m = credit_percentage / (1 - (1 + credit_percentage) ** (-number_of_credit_periods))
            credits_matrix_payment = tuple([0] * i1 + [rate_m * credit_funds] * number_of_credit_periods)
            return credits_matrix_payment, sum(credits_matrix_payment)

        def find_body_balance(credit_funds: (int, float), body1: (int, float), number_of_credit_periods: int, i1):
            a1 = [0] * i1
            for i2 in range(number_of_credit_periods):
                a1 += [credit_funds - body1]
                credit_funds -= body1
            a1 = tuple(a1)
            return a1

        for i in range(self.__count_years):
            a = percent_balance(self.__credit_funds[i], self.__credit_percentage[i], self.__number_of_credit_periods[i],
                                i)
            self.__credits_matrix_payments += [a[0]]
            self.__credits_matrix_summa += [a[1]]
            self.__credits_matrix_overprice += [self.__credits_matrix_summa[i] - self.__credit_funds[i]]
            self.__credits_matrix_percentage_per_year += [
                self.__credits_matrix_overprice[i] / self.__number_of_credit_periods[i]]
            body = self.__credits_matrix_payments[i][-1] - self.__credits_matrix_percentage_per_year[i]
            self.__credits_matrix_balances += [find_body_balance(self.__credit_funds[i], body,
                                                                 self.__number_of_credit_periods[i], i)]
            counter_sum = 0
            for j in self.__credits_matrix_balances:
                counter_sum += j[i]
            self.__credit_balance += [counter_sum]
            if i + 1 != self.__count_years:
                self.__reinvestment += [self.__credit_balance[i]]
                if self.__reinvestment[i + 1] > self.__costs[i + 1] * (1 - self.__investments_percentage[i + 1]):
                    self.__equity_investments += [self.__costs[i + 1] - self.__reinvestment[i + 1]]
                else:
                    self.__equity_investments += [self.__costs[i + 1] * self.__investments_percentage[i + 1]]
                self.__credit_funds += [
                    self.__costs[i + 1] - self.__reinvestment[i + 1] - self.__equity_investments[i + 1]]

    def __conversion_tuple(self):
        self.__credits_matrix_payments = tuple(self.__credits_matrix_payments)
        self.__credits_matrix_summa = tuple(self.__credits_matrix_summa)
        self.__credits_matrix_overprice = tuple(self.__credits_matrix_overprice)
        self.__credits_matrix_percentage_per_year = tuple(self.__credits_matrix_percentage_per_year)
        self.__credits_matrix_balances = tuple(self.__credits_matrix_balances)
        self.__reinvestment = tuple(self.__reinvestment)
        self.__equity_investments = tuple(self.__equity_investments)
        self.__credit_funds = tuple(self.__credit_funds)
        self.__credit_balance = tuple(self.__credit_balance)

    @property
    def interests(self):
        a = []
        for i in range(self.__count_years):
            b = 0
            for j in range(i + 1):
                if self.__number_of_credit_periods[j] - 1 >= i:
                    b += self.__credits_matrix_percentage_per_year[j]
            a += [b]
        a = tuple(a)
        return a

    @property
    def amortization(self):
        a = []
        for i in range(self.__count_years):
            b = 0
            for j in range(i + 1):
                if i - j < self.__useful_life:
                    b += self.__capex[j] / self.__useful_life
            a += [b]
        a = tuple(a)
        return a

    @property
    def debt_change(self):
        a = []
        for i in range(self.__count_years):
            if i == 0:
                a += [self.__credit_balance[i]]
            else:
                a += [self.__credit_balance[i] - self.__credit_balance[i - 1]]
        a = tuple(a)
        return a

    @property
    def revenue(self):
        return self.__revenue

    @property
    def costs(self):
        return self.__costs

    @property
    def capex(self):
        return self.__capex

    @property
    def opex(self):
        return self.__opex

    @property
    def financing_equity(self):
        return self.__equity_investments

    @property
    def financing_reinvestment(self):
        return self.__reinvestment

    @property
    def financing_credit(self):
        return self.__credit_funds

    @staticmethod
    def __minus(a: tuple, b: tuple):
        c = []
        for i in range(len(a)):
            c += [a[i] - b[i]]
        c = tuple(c)
        return c

    @property
    def ebitda(self):
        return FinIndicators.__minus(self.revenue, self.opex)

    @property
    def ebit(self):
        return FinIndicators.__minus(self.ebitda, self.amortization)

    @property
    def ebt(self):
        return FinIndicators.__minus(self.ebit, self.interests)

    @property
    def taxes(self):
        tex = 0.2
        a = []
        t = False
        for i in range(len(self.ebt)):
            if not t:
                if sum(self.ebt[0:i + 1]) <= 0:
                    a += [0]
                else:
                    a += [sum(self.ebt[0:i + 1]) * tex]
                    t = True
            else:
                a += [self.ebt[i] * tex]
        a = tuple(a)
        return a

    @property
    def net_profit(self):
        return FinIndicators.__minus(self.ebt, self.taxes)

    @property
    def change_working_cap(self):
        return self.__change_working_cap

    @property
    def fcff(self):
        tex = 1 - 0.2
        a = []
        for i in range(len(self.ebit)):
            a += [self.ebit[i] * tex + self.amortization[i] + self.change_working_cap[i] - self.capex[i]]
        a = tuple(a)
        return a

    @staticmethod
    def __cumulative(a: tuple):
        c = []
        for i in range(len(a)):
            if i == 0:
                c += [a[i]]
            else:
                c += [a[i] + c[i - 1]]
        c = tuple(c)
        return c

    @property
    def cumulative_fcff(self):
        return FinIndicators.__cumulative(self.fcff)

    @property
    def fcfe(self):
        tex = 1 - 0.2
        a = []
        for i in range(len(self.fcff)):
            a += [self.fcff[i] - self.interests[i] * tex + self.debt_change[i]]
        a = tuple(a)
        return a

    @property
    def cumulative_fcfe(self):
        return FinIndicators.__cumulative(self.fcfe)

    @staticmethod
    def __discounting(a: tuple, stavka: float):
        c = []
        for i in range(len(a)):
            c += [a[i] / ((1 + stavka) ** i)]
        c = tuple(c)
        return c

    @property
    def dcf_fcff(self):
        return FinIndicators.__discounting(self.fcff, self.__discount_rate_fcff)

    @property
    def cumulative_dcf_fcff(self):
        return FinIndicators.__cumulative(self.dcf_fcff)

    @property
    def dcf_fcfe(self):
        return FinIndicators.__discounting(self.fcfe, self.__discount_rate_fcfe)

    @property
    def cumulative_dcf_fcfe(self):
        return FinIndicators.__cumulative(self.dcf_fcfe)

    @staticmethod
    def __calculate_indicators(discount_rate: float, cumulative_fcf: tuple, cumulative_dcf_fcf: tuple, fcf: tuple,
                               financing_equity):
        # интересный момент с tuple почему-то пишет, что  financing_equity это list
        dr = discount_rate
        for i in range(len(cumulative_fcf)):
            if cumulative_fcf[i] > 0:
                pp = i + 1
                break
        else:
            pp = 'не окупается'
        for i in range(len(cumulative_dcf_fcf)):
            if cumulative_fcf[i] > 0:
                dpp = i + 1
                break
        else:
            dpp = 'не окупается'
        npv = cumulative_dcf_fcf[-1]
        irr = np_f.irr(fcf)
        if math.isnan(irr):
            irr = 'Невозможно посчитать'
        pi = npv / sum(FinIndicators.__discounting(financing_equity, discount_rate))
        a = dr, pp, dpp, irr, pi
        return a

    @property
    def indicators_fcff(self):
        return FinIndicators.__calculate_indicators(self.__discount_rate_fcff, self.cumulative_fcff,
                                                    self.cumulative_dcf_fcff, self.fcff, self.financing_equity)

    @property
    def indicators_fcfe(self):
        return FinIndicators.__calculate_indicators(self.__discount_rate_fcfe, self.cumulative_fcfe,
                                                    self.cumulative_dcf_fcfe, self.fcff, self.financing_equity)
