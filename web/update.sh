# Copy files to server

SERVER=myrtle.kent.ac.uk
TARGET_DIR=webpages/wyc/web
FILES="banner.png instructions.html script.js styles.css sv-modal.css sv-modal.js sv-modal.scss xl2cal.html xlsx.mini.min.js"

echo rsync -avuh -e ssh $FILES $SERVER:$TARGET_DIR
rsync -avuh -e ssh $FILES $SERVER:$TARGET_DIR
