import sys
import re

mamp = sys.argv[1]
wp_theme = sys.argv[2]
project_root = sys.argv[3]

# extrag versiune wordpress
with open("{0}/htdocs/{1}/wp-includes/version.php".format(mamp, wp_theme), "r") as file:
    f = file.read()

regexUrl = re.compile('''
        # $wp_version = '4.9.7';
        (\$wp_version[ ]*=[ ]*'|")    #start part
        ([0-9.]+)
        ([']|")         # end part
        ''', re.VERBOSE)
url = regexUrl.search(f)
wp_ver = url.group(2)

# o scriu in fisier destinatie
with open("{}/wordpress_version.txt".format(project_root), "w") as g:
    g.write(wp_ver)