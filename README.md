# Module 2 Challenge - VBA

## The Purpose:

  ### The purpose of the project was to compare runtimes for two VBA codes
  ### using the same dataset to display the same results.

## The Results:

  ### Refactored Code

    #### The refactored code looped through 3013 rows **ONCE** to gather the data.
    #### The refactored code stored data in arrays before displaying outcomes.

    #### 2017: 0.2265625 seconds runtime.
    #### 2018: 0.140625 seconds runtime.

  ### Original Code

  #### The original code looped through 3013 rows twelve times to gather the data.
  #### The original code output data for each ticker before looping through entire dataset again.

  #### 2017: 0.6132813 seconds runtime.
  #### 2018: 0.6367188 seconds runtime.

## The Summary:

  ### Advantage for the refactored code:
    #### Faster runtimes.
  ### Disadvantage for the refactored code:
    #### Time spent writing code a second time.

  ### Advantage for the original code:
    #### Only had to be written once.
  ### Disadvantage for the original code:
    #### Slower runtimes.
