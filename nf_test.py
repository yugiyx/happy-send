import edfa
import instrument

# Select Test PD ID
test_pd = 'NF'
test_edfa = 'E1'

# Config test enviorment
edfa = edfa.EdfaOFP2(2)
att_pm = instrument.AttPmN7752A(30)
switch1 = instrument.SwZHDIY(6)
switch2 = instrument.SwZ4400(48)
osa = instrument.OsaAQ6370(20)
file = instrument.DataProcess()

# Initial test module.Modified according to different modules
osa_a_offset = 0.98
osa_b_offset = 1.64
switch1.set_sw('MLS_EDFA_PM')
att_pm.enable_apc_mode(1)
att_pm.set_apc_value(-10)
edfa.set_edfa_mode(test_edfa, 'AGC')
edfa.write_reg('41', '11')  # Modified
edfa.set_edfa_gain(test_edfa, 18)

# Read set values from excel file
all_set_values = file.open_config('test_case.xls', test_pd)
print(all_set_values)
pin_set_values = all_set_values[0]
gain_set_values = all_set_values[1]
tilt_set_values = all_set_values[2]
il_set_values = all_set_values[3]

# Test start
case_num = range(len(pin_set_values))
for x in case_num:
    print('case', x + 1, 'start')
    # Set gain
    edfa.set_edfa_gain(test_edfa, gain_set_values[x])
    print(gain_set_values[x])
    # Set tilt
    tilt_set_value = hex(int(10 * tilt_set_values[x]))[2:] \
        if tilt_set_values[x] >= 0 \
        else hex(int(65536 + 10 * tilt_set_values[x]))[4:]
    edfa.write_reg('48', tilt_set_value)  # Modified
    print(tilt_set_value)
    # Set middle stage insert loss
    if il_set_values[x] == 0:
        switch2.set_sw(23, 0)
        switch2.set_sw(24, 0)
    elif il_set_values[x] == 1:
        switch2.set_sw(23, 1)
        switch2.set_sw(24, 1)
    else:
        print('il_set_values is not 0 or 1')
    print(il_set_values[x])
    # Set the pin
    att_pm.set_apc_value(pin_set_values[x])
    # Scan trace A
    switch1.set_sw('MLS_OSA')
    a_data = osa.sweep_trace('A', osa_a_offset)
    print('Scan trace A')
    # Scan trace B
    switch1.set_sw('MLS_EDFA_OSA')
    b_data = osa.sweep_trace('B', osa_b_offset)
    print('Scan trace B')
    # Run OSA EDFA-NF function
    nf_data = osa.analysis_nf()
    print('Get EDFA NF data')
    # Save test data to excel file
    save_name = '25_' + str(x + 1) + '_pin' + str(pin_set_values[x]) + '_gain'\
                + str(gain_set_values[x]) + '_tilt' + str(tilt_set_values[x])\
                + '_IL' + str(il_set_values[x] * 5) + '.xls'
    print(save_name)
    file.save_nf_data(save_name, a_data, b_data, nf_data)
