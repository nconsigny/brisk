# Function & method from prEN 1991-1-2:2021(E) & prEN 1995-1-2 v10 2024-04-04

######
# EN 1991-1-2 Annex A (informative) Parametric temperature-time curves
"""
A.2 Scope and field of application
(1) The temperature-time curves given in this Annex are valid for fire compartments up to 500 m2 of floor area,
without openings in the roof and for a maximum compartment height of 4 m.
It is assumed that the fire load of the compartment is completely burnt out.
(2) If fire load densities are specified
without specific consideration to the combustion behaviour (see Annex E), then this approach should be limited to
fire compartments with mainly cellulosic type fire loads.
"""

from sympy import exp, sqrt

fire_growth = 'medium'


def heating_phase_temperature(t, gamma) -> float:
    # gamma = (o / b) ** 2 / (0.04 / 1160) ** 2  # []
    t_star = t * gamma  # [h]
    # e1 = 0.324 * exp(-0.2 * t_star)
    # e2 = 0.203 * exp(-1.7 * t_star)
    # e3 = 0.472 * exp(-19 * t_star)
    r = 20 + 1325 * (1 - 0.324 * exp(-0.2 * t_star) - 0.203 * exp(-1.7 * t_star) - 0.472 * exp(-19 * t_star))
    return round(float(r), 0)  # [°C]


def cooling_phase_temperature(t, gamma, tm_star, temp_max, x) -> float:
    # gamma = (o / b) ** 2 / (0.04 / 1160) ** 2  # []
    t_star = t * gamma  # [h]
    if tm_star <= 0.5:
        r = max(temp_max - 625 * (t_star - tm_star * x), 20)
    elif tm_star < 2:
        r = max(temp_max - 250 * (3 - tm_star) * (t_star - tm_star * x), 20)
    else:
        r = max(temp_max - 250 * (t_star - tm_star * x), 20)
    return round(r, 0)


def cooling_phase_end(gamma, tm_star, temp_max, x) -> float:
    if tm_star <= 0.5:
        return (temp_max + 625 * tm_star * x - 20) / (625 * gamma)
    elif tm_star < 2:
        r = -(temp_max - 250 * tm_star ** 2 * x + 750 * tm_star * x - 20) / (250 * gamma * (tm_star - 3))
        return r
    else:
        return (temp_max + 250 * tm_star * x - 20) / (250 * gamma)


def thermal_absorptivity(ro, c, _lambda):
    b = sqrt(ro * c * _lambda)  # [J/m2s1/2K]
    return [b, (100 <= b <= 2200)]


def set_fire_growth(s):
    global fire_growth
    fire_growth = s


def t_lim():
    if fire_growth == 'slow':
        return 25
    elif fire_growth == 'medium':
        return 20
    elif fire_growth == 'fast':
        return 15


def t_max(q_td, o):  # [mn]
    return max(60 * 0.2e-3 * q_td / o, t_lim())


def s_lim(ro, c, _lambda, tm):  # (A.4)
    return sqrt((3600 * tm * _lambda) / (ro * c))


# EC5-1-2: 5.4.2.2 Notional design charring rate
def beta_0():
    # TODO other cases of Table 5.4
    return 0.65


def beta_n(linear, circular=False):  # [mm] (5.2)
    # Table 5.3 — Modification factors ki for charring
    k_con = 1  # 5.4.2.2(3) and (4) ! IGNORED (considered as local)
    k_gd = 1  # 5.4.2.2(5) and (6) ! IGNORED (considered as local)
    k_g = 1  # 7.2.3(2) and (3) ! IGNORED (considered as local)
    k_h = 1  # 5.4.2.2(7) ! IGNORED (For solid wood panelling and cladding's, LVL panels and wood-based panels)
    if linear:
        if circular:
            k_n = 1.3
        else:
            k_n = 1.08
    else:
        k_n = 1  # 7.2.2(2) and (3) ! ST & FST IGNORED
    # TODO combined section factor ! IGNORED for that version (considered as local) 7.2.4(14)
    k_sn1 = 1
    k_sn2 = 1
    kp = 1  # 5.4.2.2(8) ! IGNORED (no wood-based panels: kp=sqrt(450/ro) )
    # TODO Protections factors (k2 to k4) ! IGNORED for that version
    k2 = 1
    k3 = 1
    k3_1 = 1
    k3_2 = 1
    k4 = 1
    return k_con * k_gd * k_g * k_h * k_n * k_sn1 * k_sn2 * kp * k2 * k3 * k3_1 * k3_2 * k4 * beta_0()  # [mm]
