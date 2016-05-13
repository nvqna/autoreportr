import sys

if len(sys.argv) != 2:
    print "Usage: python %s risk_register.xlsx" % sys.argv[0]
    sys.exit(0)

risk_register = sys.argv[1]

print "Reading from risk_register"


