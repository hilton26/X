# Source: https://towardsdatascience.com/do-stuff-at-each-ode-integration-step-monkey-patching-solve-ivp-359b39d5f2from functools import partial

import matplotlib.pyplot as plt
import numpy as np
from scipy.integrate import solve_ivp

#
# <--- Here's where you'll put the monkey patch we'll see in a second
#

# dydt function for the Lorenz Equations
def lorenz_dydt(_t, y, s=10, r=28, b=2.667):

    xp = s*(y[1] - y[0])
    yp = y[0] * (r - y[2]) - y[1]
    zp = y[0]*y[1] - b*y[2]

    return np.asarray([xp, yp, zp])


if __name__ == '__main__':

    # parameters
    sigma = 10
    rho = 28
    beta = 2.667

    # fix the parameters
    lorenz_dydt_ode = partial(lorenz_dydt, s=sigma, r=rho, b=beta)

    # solve the system
    y0 = np.asarray([.2, .3, .4])
    sol = solve_ivp(lorenz_dydt_ode, [0, 40], y0, method='RK45', rtol=1e-12)

    # optional plotting
    ax = plt.axes(projection='3d')
    ax.plot3D(*[column for column in sol.y], 'blue')
    ax.set_title(fr'Lorenz Equations ($\sigma={sigma}, \rho={rho}, \beta={beta}$)')
    plt.show()