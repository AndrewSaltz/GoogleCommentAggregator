# Google Comment Aggregator

I'm a teacher, and I require my students to comment on other's documents. I need some way to quickly assess who did 
and did not do this.
Google Comment Aggregator grabs all comments from a group of Google Docs and places them on a spreadsheet, 
with user name and date added. It ignores empty comments and "reopened". It cannot read closed comments. It probably
could be more pythonic, as I'm new to all this.


### Prerequisites

Python 3

Git, or some other way to get this from GitHub


### Installing

If you are familiar with this pip, python, and GitHub:
```
git clone https://github.com/AndrewSaltz/GoogleCommentAggregator.git comments
cd comments
pip install -r requirements.txt
```
If not, let me recommend reading this [piece](https://readwrite.com/2013/09/30/understanding-github-a-journey-for-beginners-part-1/)
If you are on a Mac (hi, Philly teachers), also read [this](https://github.com/nicolashery/mac-dev-setup)


## What to do

1. Download you Google Docs as a zip. This is automatic if you select a bunch of Google Documents to download
2. ``` python get_comments.py ```
3. Select the zip you downloaded in step 1
4. That's it! Your spreadsheet will be in the same folder as get_comments.py

## Authors

* **Andrew Saltz** - *Initial work* - [AndrewSaltz](https://github.com/AndrewSaltz), also on [Twitter](https://twitter.com/mr_saltz)

## License

This project is licensed under the MIT License

## Acknowledgments

*Shout out to our lord and savior Stack Overflow
