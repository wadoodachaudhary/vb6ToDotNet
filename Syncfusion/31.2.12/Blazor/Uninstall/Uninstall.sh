#!/bin/sh

#  Uninstall.sh
#
#
#  Created by labuser on 20/09/19.
#

# Delete the installed files location...

rm -r $HOME/Documents/Syncfusion/$PRODUCT_VERSION/Blazor

FILE_VERSION=""
DIR_VERSION="$HOME/Documents/Syncfusion/$PRODUCT_VERSION"
# init
# look for empty dir
if [ "$(ls -A $DIR_VERSION)" ]; then
echo ""
else
rm -r $DIR_VERSION
fi

FILE_SYNCFUSION=""
DIR_SYNCFUSION="$HOME/Documents/Syncfusion"
# init
# look for empty dir
if [ "$(ls -A $DIR_SYNCFUSION)" ]; then
echo "Take action $DIR is not Empty"
else
rm -r $DIR_SYNCFUSION
fi

# Delete the System Entries....

sudo pkgutil --forget com.synfusion.syncfusionEssentialStudioForBlazor.ReadMe_$PRODUCT_VERSION.pkg
sudo pkgutil --forget com.synfusion.syncfusionEssentialStudioForBlazor.Uninstall_$PRODUCT_VERSION.pkg
sudo pkgutil --forget com.synfusion.syncfusionEssentialStudioForBlazor.LicenseAgreement_$PRODUCT_VERSION.pkg
sudo pkgutil --forget com.synfusion.syncfusionEssentialStudioForBlazor.ReleaseNotes_$PRODUCT_VERSION.pkg
sudo pkgutil --forget com.synfusion.syncfusionEssentialStudioForBlazor.Samples_$PRODUCT_VERSION.pkg
sudo pkgutil --forget com.synfusion.syncfusionEssentialStudioForBlazor.Scripts_$PRODUCT_VERSION.pkg
sudo pkgutil --forget com.synfusion.syncfusionEssentialStudioForBlazor.Source_$PRODUCT_VERSION.pkg
