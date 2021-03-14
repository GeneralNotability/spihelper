import argparse
import pywikibot


def main(target, file_to_deploy, revision_id):
    site = pywikibot.Site('en', 'wikipedia')
    page = pywikibot.Page(site, target)
    with open(file_to_deploy, 'r') as source_file:
        page.text = source_file.read()

    page.save(f'Deploying revision {revision_id}')

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--target', help='Target page')
    parser.add_argument('--file', help='File to deploy')
    parser.add_argument('--revision-id', help='Revision ID')
    args = parser.parse_args()
    main(args.target, args.file, args.revision_id)
