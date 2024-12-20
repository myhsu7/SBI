#!/bin/bash

var_no_spaces=$(echo "$1" | tr -d ' ')
echo "Removing space in file name: renaming $1 to $var_no_spaces"
echo ""
echo "new filename: $var_no_spaces \n"
echo "rename 's/ //g' '$1'"
rename 's/ //g' "$1"

# remove silence can reduce hallucination from whisper model
start_time=$2
end_time=$3

echo ""

rm $var_no_spaces.wav
# set highpass and lowpass to keep vocal voice only
ffmpeg -i $var_no_spaces -ar 16000 -ac 1 -ss $start_time -to $end_time -c:a pcm_s16le \
   -af "volume=2, highpass=f=200, lowpass=f=3000" $var_no_spaces.wav

echo ""

./main -m ./models/ggml-large-v2.bin -l chinese -f $var_no_spaces.wav -otxt
