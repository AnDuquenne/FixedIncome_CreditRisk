import xlwings as xw
import numpy as np

from scipy.stats import norm

def print_in_file(s):
    """
    This function prints a message in the output.txt file
    :return:
    """
    with open("output.txt", "a") as f:
        f.write(s + "\n")


def bootstrap_swap_rate():
    wb = xw.Book.caller()
    sheet = wb.sheets["IRS - Interest Rate Swap"]
    N = float(sheet.range('D19').value)
    n = float(sheet.range('D20').value)
    delta = float(sheet.range('D21').value)
    r = float(sheet.range('D22').value)
    last_i = int(sheet.range('D23').value)

    c_ = np.array(sheet.range(f'H21:W21').value)
    # remove 0 at then end of the array
    c_ = c_[:last_i]

    Z_0 = (1 + (r / n)) ** (-n*delta)

    # Display the value of Z_0
    sheet.range('F25').value = f"Z(0,{delta})"
    sheet.range('G25').value = Z_0

    # Calculate the value of Z_i
    Z_i = np.zeros(last_i)
    Z_i[0] = Z_0

    for j in range(1, last_i):
        Z_i[j] = (1 - (c_[j] * delta * np.sum(Z_i[:j]))) / (1 + (c_[j] * delta))

    # Display the value of Z_i
    for i in range(last_i):
        sheet.range(f'J{25 + i}').value = f"Z({0}, {i*delta+delta})"
        sheet.range(f'K{25 + i}').value = Z_i[i]

def inverse_matrix_Z_from_obs_mark_prices():
    wb = xw.Book.caller()
    sheet = wb.sheets["Bond pricing"]

    matrix = np.array(sheet.range('N64:R68').value)
    matrix_size_before = matrix.shape

    # Remove rows with all zeros
    matrix = matrix[~np.all(matrix == 0, axis=1)]

    # Remove columns with all zeros
    matrix = matrix[:, ~np.all(matrix == 0, axis=0)]

    # Check if the matrix is square
    if matrix.shape[0] != matrix.shape[1]:
        print_in_file("The matrix is not square")
        return

    matrix_size_after = matrix.shape

    starting_cell = "N71"
    ending_cell = chr(ord('N') + matrix_size_after[1] - 1) + str(73 + matrix_size_after[0])

    inv_mat = np.linalg.inv(matrix)
    sheet.range(f'{starting_cell}:{ending_cell}').value = inv_mat

    print_in_file(f"{sheet.range('N64:R68').value}")
    print_in_file(f"{type(sheet.range('N74').value)}")

    P = np.array(sheet.range('L64:L68').value).T

    # remove 0 at then end of the array
    P = P[:matrix_size_after[0]]

    Z = inv_mat @ P

    print_in_file(f"{Z}")

    for i in range(matrix_size_after[0]):
        sheet.range(f'X{64 + i}').value = Z[i]

def floor_pricing():

    wb = xw.Book.caller()
    sheet = wb.sheets["Caps & Floors"]

    f_n_matrix = np.array(sheet.range('AC40:AQ54').value, dtype=np.float64)
    diag = np.diag(f_n_matrix)

    t = float(sheet.range('H11').value)
    N = float(sheet.range('D7').value)
    n = int(sheet.range('D8').value)
    delta = float(sheet.range('D9').value)
    sigma = float(sheet.range('D10').value)
    k = float(sheet.range('D11').value)
    last_i = int(sheet.range('D12').value)
    nb_floorlets = int(sheet.range('D13').value)
    T1 = float(sheet.range('D17').value)
    T2 = float(sheet.range('D18').value)
    T = np.array(sheet.range('I59:W59').value, dtype=np.float64)
    T_i_1 = np.array(sheet.range('I65:W65').value, dtype=np.float64)
    Z = np.array(sheet.range('I68:W68').value, dtype=np.float64)

    # Display values in excel
    starting = "I"
    for i in range(diag.shape[0]):
        # Display the value horizontally
        sheet.range(f'{chr(ord(starting) + i)}69').value = diag[i]

    # Compute d1 and d2
    d1 = (np.log(diag / k) + ((sigma ** 2 / 2) * (T_i_1 - t))) / (sigma * np.sqrt(T_i_1 - t))
    d2 = d1 - sigma * np.sqrt(T_i_1 - t)

    # Phi(d1) and Phi(d2)
    phi_d1 = np.array([norm.cdf(d1[i]) for i in range(d1.shape[0])])
    phi_d2 = np.array([norm.cdf(d2[i]) for i in range(d2.shape[0])])

    d1[nb_floorlets:] = 0
    d2[nb_floorlets:] = 0
    phi_d1[nb_floorlets:] = 0
    phi_d2[nb_floorlets:] = 0

    # Compute the value of the floor and the cap
    caplets = np.zeros(d1.shape)
    floorlets = np.zeros(d1.shape)
    caplets = N * delta * Z * (diag * phi_d1 - k * phi_d2)
    floorlets = caplets - N * delta * Z * (diag - k)

    # Display values in excel
    starting = "I"
    for i in range(phi_d1.shape[0]):
        # Display the value of Z_i horizontally
        sheet.range(f'{chr(ord(starting) + i)}71').value = d1[i]
        sheet.range(f'{chr(ord(starting) + i)}72').value = d2[i]
        sheet.range(f'{chr(ord(starting) + i)}73').value = phi_d1[i]
        sheet.range(f'{chr(ord(starting) + i)}74').value = phi_d2[i]
        sheet.range(f'{chr(ord(starting) + i)}76').value = caplets[i]
        sheet.range(f'{chr(ord(starting) + i)}77').value = floorlets[i]

    floor = np.sum(floorlets)
    cap = np.sum(caplets)

    sheet.range('G79').value = floor
    sheet.range('G80').value = cap