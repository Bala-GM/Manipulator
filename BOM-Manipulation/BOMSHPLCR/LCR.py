'''import re  # CAP

component_lines = [
        "SMT,H/IND 1.2uH 8.6A 4.3X4.3X2.0 20% BOURNS",
        "FERRITE BEAD SMD 1.8K OHM 0402 1LN",
        "Capacitor,SMT,1206,10uF,16V,10%,Tantalum",
        "Cap SMD Aluminum Lytic 4700uF 16V 20% (18 X 21.5mm) 1750mA 2000h 85°C Automotive T/R",
        "Cap Super  THT 3.6-3.69V Impedance:≤ 500 Mω 21.5x9mm",
        "CAP,MLCC,SMD,220PF,50V,±5%,0402,C0G,(NP0)",
        "RES SMD 3.3M OHM 1% 1/10W 0603",
        "RES,THICK,FILM,SMD,47K,OHM,1%,0.0625W,0402,100ppm",
        "EMI,filter,bead,SMT,0402,1k,ohms,Tape,GHz,Band,Gen,Use",
        "FERRITE BEAD SMD 1K OHM 0603 1LN",
        "Ferrite,Chip,SMT,0805,Filter,330R",
        "IND CMC 2A 2LN 500 OHM SMD",
        "ALLUM CAP SMD 100uF 6.3volts 20 % AEC-Q200",
        "ALUM CAP SMD 100UF 6.3V 20% (6.60mm x 6.60mm)",
        "Aluminum Electrolytic Capacitors - SMD 1500uF 35V",
        "CAP ALUM 1500UF 20% 35V SMD",
        "CAP TAN  SMD  4.7UF  10V  ±20% 0805",
        "CAP ALU  SMD 330uF ± 20% 35V",
        "A/CAP 2.2uF 0402 6.3V K(+-10%) X5R",
        "A/CAP 4.7uF 0402 ≧6.3V ≦M X5R",
        "SMT,A/CAP 10uF 0603 ≧25V ≦M X5R",
        "SMT,A/CAP 22uF 0805 ≧25V ≦M X5R",
        "H/CAP 47uF 0603 ≧6.3V ≦M X5R",
        "H/RES 15Ω 0402 ≦J 1/16W",
        "H/RES 47Ω 0402 ≦J 1/16W",
        "H/RES 240Ω 0201 F(+-1%) 1/20W",
        "H/RES 240Ω 0402 ≦F 1/16W",
        "SMT,H/RES 249Ω 0603 F(+-1%) 1/10W"
]

for component_line in component_lines:
    capacitance_match = re.search(r'(\d+(\.\d+)?)([pnuμmkKM]F)', component_line)
    tolerance_match = re.search(r'[± ](\d+)%', component_line)

    if capacitance_match:
        capacitance_value, _, capacitance_unit = capacitance_match.groups()
        print(f"Capacitance Value: {capacitance_value}, Unit: {capacitance_unit}")

    if tolerance_match:
        tolerance_value = tolerance_match.group(1)
        print(f"Tolerance Value: ±{tolerance_value}%")

    print()'''


'''import re #RES

component_line = "H/RES 15Ω 0402 ≦J 1/16W"

# Regular expressions for RESISTOR
resistor_type_match = re.search(r'(RES)', component_line)
resistor_value_match = re.search(r'(\d+(\.\d+)?)([pnuμmkKMΩ]?)', component_line)
tolerance_match = re.search(r'(\d+)%', component_line)

if resistor_type_match:
    resistor_type = resistor_type_match.group(1)
    print(f"Resistor Type: {resistor_type}")

if resistor_value_match:
    resistor_value, _, resistor_unit = resistor_value_match.groups()
    print(f"Resistor Value: {resistor_value}{resistor_unit}")

if tolerance_match:
    tolerance_value = tolerance_match.group(1)
    print(f"Tolerance: {tolerance_value}%")'''



'''import re

resistor_lines = [
"SMT,H/RES 249Ω 0603 F(+-1%) 1/10W",
"RES SMD 560 OHM 1% 0.1W 0603",
"Thick,Film,Resistors,-,SMD,10,OHMS,1%,0.1W,0603,100PPM",
"Thick,Film,Resistors,-,SMD,15,kOhms,100-200,mW,0603,1%",
"CAP,MLCC,SMD,220PF,50V,±5%,0402,C0G,(NP0)",
"SMT,A/CAP 10uF 0603 ≧25V ≦M X5R",
"SMT,A/CAP 22uF 0805 ≧25V ≦M X5R",
"H/CAP 47uF 0603 ≧6.3V ≦M X5R",
"H/RES 15Ω 0402 ≦J 1/16W",
"H/RES 47Ω 0402 ≦J 1/16W",
]

for component_line in resistor_lines:
    # Regular expressions for RESISTOR
    resistor_type_match = re.search(r'\b(Res|Resistor|RESISTOR)\b', component_line, re.IGNORECASE)
    resistor_value_match = re.search(r'(\d+(\.\d+)?)\s*([kKMΩ]?)\s*(\d+%)?', component_line)

    if resistor_type_match:
        resistor_type = resistor_type_match.group(1)
        print(f"Resistor Type: {resistor_type}")

    if resistor_value_match:
        resistor_value, _, resistor_unit, tolerance = resistor_value_match.groups()
        print(f"Resistor Value: {resistor_value}{resistor_unit}")
        if tolerance:
            print(f"Tolerance: ±{tolerance}")

    print("---")'''


'''import re

component_lines = [
    "H/CAP 47uF 0603 ≧6.3V ≦M X5R",
    "H/RES 15Ω 0402 ≦J 1/16W",
    # Add more lines if needed
]

for component_line in component_lines:
    # Regular expressions for RESISTOR and CAPACITOR
    res_match = re.search(r'\b(Res|Resistor|RESISTOR)\b', component_line, re.IGNORECASE)
    cap_match = re.search(r'\b(Cap|Capacitor|CAPACITOR)\b', component_line, re.IGNORECASE)

    if res_match:
        print("Extracted RES:", component_line)

    if cap_match:
        print("Extracted CAP:", component_line)'''


'''import pandas as pd
import re

def extract_component_info(component_line):
    # Regular expressions for RESISTOR
    resistor_type_match = re.search(r'(RES)', component_line)
    resistor_value_match = re.search(r'(\d+(\.\d+)?)([pnuμmkKMΩ]?)', component_line)
    tolerance_match = re.search(r'(\d+)%', component_line)

    # Assign default values
    A = None
    B = None
    C = None
    D = None

    if resistor_type_match:
        A = resistor_type_match.group(1)

    if resistor_value_match:
        B = resistor_value_match.group(1)
        C = resistor_value_match.group(3)

    if tolerance_match:
        D = tolerance_match.group(1)

    return A, B, C, D

component_lines = [
    "SMT,H/RES 249Ω 0603 F(+-1%) 1/10W",
    "RES SMD 560 OHM 1% 0.1W 0603",
    "Thick,Film,Resistors,-,SMD,10,OHMS,1%,0.1W,0603,100PPM",
    "Thick,Film,Resistors,-,SMD,15,kOhms,100-200,mW,0603,1%",
    "CAP,MLCC,SMD,220PF,50V,±5%,0402,C0G,(NP0)",
    "SMT,A/CAP 10uF 0603 ≧25V ≦M X5R",
    "SMT,A/CAP 22uF 0805 ≧25V ≦M X5R",
    "H/CAP 47uF 0603 ≧6.3V ≦M X5R",
    "H/RES 15Ω 0402 ≦J 1/16W",
    "H/RES 47Ω 0402 ≦J 1/16W",
]

data = {'LCR Type': [], 'LCR Value': [], 'LCR Unit': [], 'LCR Tolerance': []}

for component_line in component_lines:
    A, B, C, D = extract_component_info(component_line)
    data['LCR Type'].append(A)
    data['LCR Value'].append(B)
    data['LCR Unit'].append(C)
    data['LCR Tolerance'].append(D)

df = pd.DataFrame(data)
print(df)
'''


import pandas as pd
import re

def extract_component_info(component_line):
    # Regular expressions for CAPACITOR
    capacitor_type_match = re.search(r'\b(CAP)\b', component_line, re.IGNORECASE)
    capacitor_value_match = re.search(r'(\d+(\.\d+)?)([pnuμmkKMΩ]?)\s*(\d+%)?', component_line)

    # Regular expressions for RESISTOR
    resistor_type_match = re.search(r'\b(Res|Resistor|RESISTOR)\b', component_line, re.IGNORECASE)
    resistor_value_match = re.search(r'(\d+(\.\d+)?)([pnuμmkKMΩ]?)\s*(\d+%)?', component_line)

    # Assign default values
    LCR_Type = None
    LCR_Value = None
    LCR_Unit = None
    LCR_Tolerance = None

    if capacitor_type_match:
        LCR_Type = capacitor_type_match.group(1)

    if capacitor_value_match:
        groups = capacitor_value_match.groups()
        LCR_Value = groups[0] if groups[0] is not None else None
        LCR_Unit = groups[2] if groups[2] is not None else None

    if resistor_type_match:
        LCR_Type = resistor_type_match.group(1)

    if resistor_value_match:
        groups = resistor_value_match.groups()
        LCR_Value = groups[0] if groups[0] is not None else None
        LCR_Unit = groups[2] if groups[2] is not None else None

    return LCR_Type, LCR_Value, LCR_Unit, LCR_Tolerance

component_lines = [
    "SMT,H/RES 249Ω 0603 F(+-1%) 1/10W",
    "RES SMD 560 OHM 1% 0.1W 0603",
    "Thick,Film,Resistors,-,SMD,10,OHMS,1%,0.1W,0603,100PPM",
    "Thick,Film,Resistors,-,SMD,15,kOhms,100-200,mW,0603,1%",
    "CAP,MLCC,SMD,220PF,50V,±5%,0402,C0G,(NP0)",
    "SMT,A/CAP 10uF 0603 ≧25V ≦M X5R",
    "SMT,A/CAP 22uF 0805 ≧25V ≦M X5R",
    "H/CAP 47uF 0603 ≧6.3V ≦M X5R",
    "H/RES 15Ω 0402 ≦J 1/16W",
    "H/RES 47Ω 0402 ≦J 1/16W",
]

data = {'LCR Type': [], 'LCR Value': [], 'LCR Unit': [], 'LCR Tolerance': []}

for component_line in component_lines:
    LCR_Type, LCR_Value, LCR_Unit, LCR_Tolerance = extract_component_info(component_line)
    data['LCR Type'].append(LCR_Type)
    data['LCR Value'].append(LCR_Value)
    data['LCR Unit'].append(LCR_Unit)
    data['LCR Tolerance'].append(LCR_Tolerance)

df = pd.DataFrame(data)
print(df)
