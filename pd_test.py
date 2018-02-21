import time
import edfa
import instrument

# Select Test PD ID
test_pd = 'APC1'
test_edfa = 'E1'

# Config test enviorment
tls_1510_pd = ['PD2']
tls_1610_pd = ['PD16']
out_pd = ['PD9', 'PDT1', 'PDT2', 'APC1', 'APC2']
edfa = edfa.EdfaOFP2(2)
att_pm = instrument.AttPmN7752A(30)
pm = instrument.Pm8163A(21)
switch = instrument.SwZHDIY(6)
file = instrument.DataProcess()

# Initial test module.Modified according to different modules
if test_pd in tls_1510_pd:
    pm.set_pm_wavlength(1510)
    att_pm.set_att_wavlength(1510)
    switch.set_sw('TLS_EDFA_PM')
    pin_offset = 0
    pout_offset = 1.8
elif test_pd in tls_1610_pd:
    pm.set_pm_wavlength(1610)
    att_pm.set_att_wavlength(1610)
    switch.set_sw('TLS_EDFA_PM')
    pin_offset = 0
    pout_offset = 1.8
else:
    pm.set_pm_wavlength(1550)
    att_pm.set_att_wavlength(1550)
    switch.set_sw('MLS_EDFA_PM')
    pin_offset = 1
    pout_offset = -1
att_pm.enable_apc_mode(1)
att_pm.set_apc_value(-10)
edfa.set_edfa_mode(test_edfa, 'AGC')
edfa.write_reg('41', '11')  # Modified
edfa.set_edfa_gain(test_edfa, 18)

# Read set values from excel file
all_set_values = file.open_config('test_case.xls', test_pd)
pd_report = [[], [], []]
pm_report = []
print(all_set_values[0])

# Test start
for set_value in all_set_values[0]:
    # Set the pin or pout
    if test_pd == 'APC1' or test_pd == 'APC2':
        edfa.set_edfa_mode(test_edfa, 'APC')
        edfa.set_edfa_power(test_edfa, set_value)
    else:
        att_pm.set_apc_value(set_value + pin_offset)
    print('Set value:', set_value)
    # Read test pd report value 3 times
    t = 0
    while t <= 2:
        read_data = edfa.read_pd(test_pd)
        pd_report[t].append(read_data)
        t += 1
        time.sleep(0.5)
        print('PD report:', read_data)
    # Read pm report value
    if test_pd in out_pd:
        read_data = pm.get_pm_value() + pout_offset
        pm_report.append(read_data)
        print('PM report:', read_data)
print(all_set_values[0])
print(pd_report[0])
print(pd_report[1])
print(pd_report[2])
print(pm_report)

# Save test data to excel file
test_data = [all_set_values[0],
             pd_report[0], pd_report[1], pd_report[2]]
if test_pd in out_pd:
    test_data.append(pm_report)
file.save_pd_data('test_result_pd.xls', test_pd, test_data)
