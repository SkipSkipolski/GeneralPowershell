#General Usage

Param(
    [parameter(Position=0,
    Mandatory=$true,
    ValueFromPipeline=$true
    ParameterSetName="Computer"
    ValueFromPipelineByPropertyName=$true
    ValueFromRemainingArguments=$true
    HelpMessage="Enter one or more computer names separated by commas.")]
    alias("CN","MachineName")
    [String[]]
    $ComputerName,

    [parameter(Position=1)]
    [AllowNull(),
    AllowEmptyString(),
    AllowEmptyCollection(),
    ValidateCount(1,5),
    ValidateLength(1,10),
    ValidatePattern("[0-9][0-9][0-9][0-9]",
    ValidateRange(0,10),
    ValidateScript({$_ -ge (Get-Date)}),
    ValidateSet("Low", "Average", "High")
    ]
)

1) Parameters
2) Functions
3) Logging
