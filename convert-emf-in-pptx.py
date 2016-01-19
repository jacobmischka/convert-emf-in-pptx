#!/usr/bin/env python3

# Requires `unoconv` and imagemagick `convert` to be installed and in path

import os, os.path, shutil, sys, fileinput, tempfile
from subprocess import run, PIPE
from zipfile import ZipFile

# Anything supported by imagemagick convert
output_format = ".png"

def main():
	if sys.argv[1]:
		pptx_path = sys.argv[1]
	else:
		return "Please pass the path to the pptx as the first argument."


	if sys.argv[2]:
		new_pptx_path = sys.argv[2]
		if ".pptx" not in new_pptx_path:
			new_pptx_path += ".pptx"
	else:
		new_pptx_path = os.path.splitext(pptx_path)[0] + "_converted.pptx"

	with tempfile.TemporaryDirectory() as tempdir:

		# Extract original pptx
		with ZipFile(pptx_path) as pptx_zip:
			pptx_zip.extractall(tempdir)

		media_dir = os.path.join(tempdir, "ppt", "media")
		original_wd = os.getcwd()

		for img in os.listdir(media_dir):
			if os.path.splitext(img)[1] == ".emf":
				# Convert to pdf with unoconv
				run(["unoconv", "-f", "pdf", os.path.join(media_dir, img)])
				img_basename = os.path.splitext(img)[0]
				img_pdf = img_basename + ".pdf"

				# EMFs are weird and I haven't actually read the spec
				# Embedded image is off center, but it seems to want the offsets

				# Find amount to be trimmed from top and left of image
				trim_info = run(["convert", os.path.join(media_dir, img_pdf),
					"-trim", "info:"], stdout=PIPE, universal_newlines=True)
				pdf_info = trim_info.stdout.split(" ")
				pdf_img_info = pdf_info[-5].split("+")
				trim_x = pdf_img_info[1]
				trim_y = pdf_img_info[2]

				img_out = img_basename + output_format

				# Trim them from both sides because that seems to work
				# Create trimmed image in new format
				run(["convert", os.path.join(media_dir, img_pdf), "-shave",
					trim_x+"x"+trim_y, os.path.join(media_dir, img_out)])

				# Replace references to original image name to new one
				os.chdir(os.path.join(tempdir, "ppt", "slides", "_rels"))
				with fileinput.FileInput(os.listdir("."), inplace=True) as rel:
					for line in rel:
						print(line.replace(img, img_out), end="")
				os.chdir(original_wd)

				# Remove original and intermediary image
				os.remove(os.path.join(media_dir, img_pdf))
				os.remove(os.path.join(media_dir, img))

		# Create new pptx file
		with ZipFile(new_pptx_path, "w") as new_pptx:
			os.chdir(tempdir)
			for f in os.scandir():
				if f.is_file():
					new_pptx.write(f.name)
				elif f.is_dir:
					for root, dirs, files in os.walk(f.name):
						for file in files:
							new_pptx.write(os.path.join(root, file))



if __name__ == '__main__':
	main()
