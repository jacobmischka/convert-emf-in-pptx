#!/usr/bin/env python3

# Requires `unoconv` and imagemagick `convert` to be installed and in path
from argparse import ArgumentParser
import os, fileinput, tempfile
from subprocess import run, PIPE
from zipfile import ZipFile

def convert(inpath, outpath, output_format, output_size, output_quality):
	with tempfile.TemporaryDirectory() as tempdir:

		# Extract original pptx
		with ZipFile(inpath) as pptx_zip:
			pptx_zip.extractall(tempdir)

		media_dir = os.path.join(tempdir, 'ppt', 'media')
		original_wd = os.getcwd()

		for img in os.listdir(media_dir):

			img_basename = os.path.splitext(img)[0]
			img_out = img_basename + output_format

			if os.path.splitext(img)[1] == '.emf':
				# Convert to pdf with unoconv
				run(['unoconv', '-f', 'pdf', os.path.join(media_dir, img)])
				img_pdf = img_basename + '.pdf'

				# EMFs are weird and I haven't actually read the spec
				# Embedded image is off center, but it seems to want the offsets

				# Find amount to be trimmed from top and left of image
				trim_info = run([
					'convert', os.path.join(media_dir, img_pdf),
					'-trim',
					'info:'
				], stdout=PIPE, universal_newlines=True)
				pdf_info = trim_info.stdout.split(' ')
				pdf_img_info = pdf_info[-5].split('+')
				trim_x = pdf_img_info[1]
				trim_y = pdf_img_info[2]

				# Trim them from both sides because that seems to work
				# Create trimmed image in new format
				run([
					'convert', os.path.join(media_dir, img_pdf),
					'-shave', trim_x + 'x' + trim_y,
					'-resize', output_size,
					'-quality', output_quality,
					os.path.join(media_dir, img_out)
				])

				# Remove original and intermediary image
				os.remove(os.path.join(media_dir, img_pdf))
				os.remove(os.path.join(media_dir, img))
			else:
				run([
					'convert', os.path.join(media_dir, img),
					'-resize', output_size,
					'-quality', output_quality,
					os.path.join(media_dir, img_out)
				])

				if img != img_out:
					os.remove(os.path.join(media_dir, img))

			# Replace references to original image name to new one
			os.chdir(os.path.join(tempdir, 'ppt', 'slides', '_rels'))
			with fileinput.FileInput(os.listdir('.'), inplace=True) as rel:
				for line in rel:
					print(line.replace(img, img_out), end='')
			os.chdir(original_wd)

		# Create new pptx file
		with ZipFile(outpath, 'w') as new_pptx:
			os.chdir(tempdir)
			for f in os.scandir():
				if f.is_file():
					new_pptx.write(f.name)
				elif f.is_dir:
					for root, _, files in os.walk(f.name):
						for file in files:
							new_pptx.write(os.path.join(root, file))

def main():
	parser = ArgumentParser(description='Shrink and convert images in .pptx presentation')
	parser.add_argument('infile', action='store', help='Presentation file to convert')
	parser.add_argument('outfile', action='store', help='Output location for new presentation')
	parser.add_argument('-f', '--format', action='store', dest='format', default='.jpg',
		help='Imagemagick-supported output file extension (default: .jpg)')
	parser.add_argument('-s', '--size', action='store', dest='size', default='1280x720<',
		help='Imagemagick-supported output image size (default: 1280x720<)')
	parser.add_argument('-q', '--quality', action='store', dest='quality', default='50',
		help='Imagemagick-supported output image quality (default: 50)')

	args = parser.parse_args()

	convert(args.infile, args.outfile, args.format, args.size, args.quality)

if __name__ == '__main__':
	main()
