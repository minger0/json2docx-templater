'''Usage: python json2docxtemplater-cli.py -h'''
from argparse import ArgumentParser, ArgumentDefaultsHelpFormatter
import json2docxtemplater as t

parser = ArgumentParser(formatter_class=ArgumentDefaultsHelpFormatter)
parser.add_argument("-C", dest="contentfolder", default="content/", help="content folder")
parser.add_argument("-T", dest="templatefolder", default="template/", help="template folder")
parser.add_argument("-O", dest="outputfolder", default="output/", help="output folder")
parser.add_argument("contentfilename", help="content file name, json format")
parser.add_argument("templatefilename", help="template file name, docx format")
args = parser.parse_args()
print(args)

templater = t.Json2docx({"contentfolder": args.contentfolder, "templatefolder": args.templatefolder, "outputfolder": args.outputfolder})
templater.fill(args.contentfilename, args.templatefilename)