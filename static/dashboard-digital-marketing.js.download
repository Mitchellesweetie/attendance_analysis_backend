$(function() {
    "use strict";

    var public_registrations = document.querySelector("#total-visitors").getAttribute('value');
    var nfoundational = document.querySelector("#direct").getAttribute('role');
    var total_reg=document.querySelector("#direct").getAttribute('name');
    var foundational = Math.round(nfoundational/total_reg*100)
    var nbasic = document.querySelector("#organic-search").getAttribute('role');
    var basic = Math.round(nbasic/total_reg*100)
    var nintemediate = document.querySelector("#referral").getAttribute('role');
    var intemediate = Math.round(nintemediate/total_reg*100)
    var nadvanced = document.querySelector("#social").getAttribute('role');
    var advanced = Math.round(nadvanced/total_reg*100)
    var regi_rate = document.querySelector("#newsletter-open-rate").getAttribute('role');
    var imported=document.querySelector("#newsletter-open-rate").getAttribute('name');
    var regist_rate = Math.round(regi_rate/imported*100)
    var cert_rate = document.querySelector("#click-through-rate").getAttribute('role');
    var certi_rate = Math.round(cert_rate/imported*100)

    var monthlydata= document.querySelector("#list-subscribers").getAttribute('role');
    // const str = "1, 67, 190, 121, 23, 22, 12, 8, 17, 26, 6, 0";

// Step 1: Convert the string to an array of numbers
const originalData = monthlydata.split(',').map(item => parseInt(item.trim(), 10));

// Step 2: Apply a transformation, e.g., adding 75 to each number (you can adjust the logic as needed)
const transformedData = originalData.map(item => item );
// var final_date
// alert(monthlydata)
// chart 1

var options = {
            chart: {
                height: 125,
                type: 'area',
                zoom: {
                      enabled: false
                    },
             foreColor: '#4e4e4e',
             toolbar: {
                  show: false
                },
          sparkline:{
              enabled:true,
            },
            dropShadow: {
                    enabled: false,
                    opacity: 0.15,
                    blur: 3,
                    left: -7,
                    top: 15,
                    //color: 'rgba(0, 158, 253, 0.65)',
                }
            },
            plotOptions: {
                bar: {
            columnWidth: '10%',
              endingShape: 'rounded',
                    dataLabels: {
                        position: 'top', // top, center, bottom
                    },
                }
            },
            dataLabels: {
                enabled: false
            },
            stroke: {
                width: 3, 
                curve: 'straight'
            },
            series: [{
                name: 'Twitter Followers',
                data: [76, 85, 101, 98, 87, 105, 91, 114, 94]
            }],

            xaxis: {
                type: 'month',
                categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"],                
            },
            yaxis: {
                axisBorder: {
                    show: false
                },
                axisTicks: {
                    show: false,
                },
                labels: {
                    show: false,
                    formatter: function(val) {
                    return parseInt(val);
                  }
                }

            },
            fill: {
                type: 'gradient',
                gradient: {
                    shade: 'light',
                    //gradientToColors: ['#6a11cb'],
                    shadeIntensity: 1,
                    type: 'vertical',
                    opacityFrom: 0.8,
                    opacityTo: 0.2,
                    stops: [0, 80, 100]
                },
            },
            //colors: ['#2575fc'],
            legend: {
                show: 0,
                position: "top",
                horizontalAlign: "center",
                offsetX: -20,
                fontSize: "12px",
                markers: {
                  radius: 50,
                  width: 10,
                  height: 10
                }
              },
            grid:{
                show: false,
                borderColor: 'rgba(66, 59, 116, 0.12)',
            },
            tooltip: {
                theme: 'dark',
                x: {
                    show: false
                },

            }
        }

        var chart = new ApexCharts(
            document.querySelector("#twitter-followers"),
            options
        );

        chart.render();
		
		
		
		
	// chart 2

    var options = {
            chart: {
                height: 290,
                type: 'radialBar',
            },
            plotOptions: {
                radialBar: {
					startAngle: -135,
                    endAngle: 225,
					hollow: {
						margin: 0,
						size: '65%',
						background: '#000',
						dropShadow: {
                        enabled: true,
                        top: -3,
                        left: 0,
                        blur: 4,
                        opacity: 0.35
                        }
					  },
                    track: {
                      background: '#fff',
                      strokeWidth: '100%',
                      margin: 0, // margin is in pixels
                      dropShadow: {
                        enabled: true,
                        top: -3,
                        left: 0,
                        blur: 4,
                        opacity: 0.35
                      }
                    },
                    dataLabels: {
                        name: {
                            fontSize: '14px',
                            color: '#fff',
                            offsetY: -10
                        },
                        value: {
                            offsetY: 0,
                            fontSize: '22px',
                            color: '#fff',
                            formatter: function (val) {
                                return val;
                            }
                        }
                    }
                },
            },
            stroke: {
                dashArray: 4
            },
            fill: {
              type: 'gradient',
              gradient: {
              shade: 'dark',
              type: 'horizontal',
              shadeIntensity: 0.5,
              gradientToColors: ['#f31368'],
              inverseColors: false,
              opacityFrom: 1,
              opacityTo: 1,
              stops: [0, 100]
            }
         },
		  colors: ["#0ce8fe"],
          series: [public_registrations],
          labels: ['Registered'],
            
        }

       var chart = new ApexCharts(
            document.querySelector("#total-visitors"),
            options
        );
        
        chart.render();

         



  // chart 3

  var options = {
            chart: {
                height: 35,
                type: 'bar',
                zoom: {
                      enabled: false
                    },
             foreColor: '#4e4e4e',
             toolbar: {
                  show: false
                },
          sparkline:{
              enabled:true,
            },
            dropShadow: {
                    enabled: true,
                    opacity: 0.1,
                    blur: 3,
                    left: -7,
                    top: 15,
                    //color: 'rgba(0, 158, 253, 0.65)',
                }
            },
            plotOptions: {
                bar: {
            columnWidth: '10%',
              endingShape: 'rounded',
                    dataLabels: {
                        position: 'top', // top, center, bottom
                    },
                }
            },
            dataLabels: {
                enabled: false
            },
            stroke: {
                width: 3, 
                curve: 'straight'
            },
            series: [{
                name: 'Facebook Pageviews',
                data: [76, 85, 101, 98, 87, 105, 91, 114, 94]
            }],

            xaxis: {
                type: 'month',
                categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"],                
            },
            yaxis: {
                axisBorder: {
                    show: false
                },
                axisTicks: {
                    show: false,
                },
                labels: {
                    show: false,
                    formatter: function(val) {
                    return parseInt(val);
                  }
                }

            },
            fill: {
                type: 'gradient',
                gradient: {
                    shade: 'light',
                    gradientToColors: ['#8f50ff'],
                    shadeIntensity: 1,
                    type: 'vertical',
                    opacityFrom: 1,
                    opacityTo: 1,
                    stops: [0, 80, 100]
                },
            },
            colors: ['#d13adf'],
            legend: {
                show: 0,
                position: "top",
                horizontalAlign: "center",
                offsetX: -20,
                fontSize: "12px",
                markers: {
                  radius: 50,
                  width: 10,
                  height: 10
                }
              },
            grid:{
                show: false,
                borderColor: 'rgba(66, 59, 116, 0.12)',
            },
            tooltip: {
                theme: 'dark',
                x: {
                    show: false
                },

            }
        }

        var chart = new ApexCharts(
            document.querySelector("#facebook-pageviews"),
            options
        );

        chart.render();
		
		
		
		

// chart 4


var options = {
            chart: {
                height: 35,
                type: 'line',
                zoom: {
                      enabled: false
                    },
             foreColor: '#4e4e4e',
             toolbar: {
                  show: false
                },
          sparkline:{
              enabled:true,
            },
            dropShadow: {
                    enabled: true,
                    opacity: 0.15,
                    blur: 3,
                    left: -7,
                    top: 5,
                    //color: 'rgba(0, 158, 253, 0.65)',
                }
            },
            plotOptions: {
                bar: {
            columnWidth: '10%',
              endingShape: 'rounded',
                    dataLabels: {
                        position: 'top', // top, center, bottom
                    },
                }
            },
            dataLabels: {
                enabled: false
            },
            stroke: {
                width: 2.5, 
                curve: 'straight'
            },
            series: [{
                name: 'Bounce Rate',
                data: [85, 101, 98, 87, 105, 91, 114, 94]
            }],

            xaxis: {
                type: 'month',
                categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"],                
            },
            yaxis: {
                axisBorder: {
                    show: false
                },
                axisTicks: {
                    show: false,
                },
                labels: {
                    show: false,
                    formatter: function(val) {
                    return parseInt(val);
                  }
                }

            },
            fill: {
                type: 'gradient',
                gradient: {
                    shade: 'light',
                    gradientToColors: ['#00b09b'],
                    shadeIntensity: 1,
                    type: 'horizontal',
                    opacityFrom: 1,
                    opacityTo: 1,
                    stops: [0, 80, 100]
                },
            },
            colors: ['#96c93d'],
            legend: {
                show: 0,
                position: "top",
                horizontalAlign: "center",
                offsetX: -20,
                fontSize: "12px",
                markers: {
                  radius: 50,
                  width: 10,
                  height: 10
                }
              },
            grid:{
                show: false,
                borderColor: 'rgba(66, 59, 116, 0.12)',
            },
            tooltip: {
                theme: 'dark',
                x: {
                    show: false
                },

            }
        }

        var chart = new ApexCharts(
            document.querySelector("#bounce-rate"),
            options
        );

        chart.render();
		



 
// chart 5

var options = {
            chart: {
                height: 200,
                type: 'area',
                zoom: {
                      enabled: false
                    },
             foreColor: '#4e4e4e',
             toolbar: {
                  show: true
                },
          sparkline:{
              enabled:false,
            },
            dropShadow: {
                    enabled: false,
                    opacity: 0.15,
                    blur: 3,
                    left: -7,
                    top: 15,
                    //color: 'rgba(0, 158, 253, 0.65)',
                }
            },
            plotOptions: {
                bar: {
            columnWidth: '30%',
              endingShape: 'rounded',
                    dataLabels: {
                        position: 'top', // top, center, bottom
                    },
                }
            },
            dataLabels: {
                enabled: false
            },
            stroke: {
                width: 0, 
                curve: 'smooth'
            },
            series: [{
                name: 'Registrations',
                data: transformedData
            }],

            xaxis: {
                type: 'month',
                categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],                
            },
            yaxis: {
                axisBorder: {
                    show: false
                },
                axisTicks: {
                    show: false,
                },
                labels: {
                    show: false,
                    formatter: function(val) {
                    return parseInt(val);
                  }
                }

            },

            fill: {
                type: 'gradient',
                gradient: {
                    shade: 'light',
                    gradientToColors: ['#ee0979'],
                    shadeIntensity: 1,
                    type: 'vertical',
                    opacityFrom: 1,
                    opacityTo: 1,
                    stops: [0, 100, 100, 100]
                },
            },
            colors: ['#ff6a00'],
            grid:{
                show: true,
                borderColor: 'rgba(66, 59, 116, 0.15)',
            },
            tooltip: {
                theme: 'dark',
                x: {
                    show: false
                },

            },
            title: {
                text: 'Monthly Registrations'
            }
        }

        var chart = new ApexCharts(
            document.querySelector("#list-subscribers"),
            options
        );

        chart.render();



     

     // chart 6

     var optionsProgress1 = {
	  chart: {
	    height: 70,
	    type: 'bar',
	    stacked: true,
	    sparkline: {
	      enabled: true
	    }
	  },
	  plotOptions: {
	    bar: {
	      horizontal: true,
	      barHeight: '20%',
	      colors: {
	        backgroundBarColors: ['#40475D']
	      }
	    },
	  },
	  dataLabels: {
             enabled: false
          },
	  stroke: {
	    width: 0,
	  },
	  series: [{
	    name: 'Foundational Digital Skills',
	    data: [foundational]
	  }],
	  title: {
	    floating: true,
	    offsetX: -10,
	    offsetY: 5,
	    text: 'Foundational Digital Skills'
	  },
	  subtitle: {
	    floating: true,
	    align: 'right',
	    offsetY: 0,
	    text: foundational+'%',
	    style: {
	      fontSize: '20px'
	    }
	  },
	  fill: {
	    type: 'gradient',
	    gradient: {
	      gradientToColors: ['#F55555']
	    },
	  },
	  colors: ['#FCCF31'],
	  tooltip: {
	  	theme: 'dark',
	  	x: {
              show: false
             },
	    enabled: true
	    
	  },
	  xaxis: {
	    categories: ['Foundational Digital Skills'],
	  },
	  yaxis: {
	    max: 100
	  }
	}

	var chartProgress1 = new ApexCharts(document.querySelector('#direct'), optionsProgress1);
	chartProgress1.render();





	// chart 7

     var optionsProgress2 = {
	  chart: {
	    height: 70,
	    type: 'bar',
	    stacked: true,
	    sparkline: {
	      enabled: true
	    }
	  },
	  plotOptions: {
	    bar: {
	      horizontal: true,
	      barHeight: '20%',
	      colors: {
	        backgroundBarColors: ['#40475D']
	      }
	    },
	  },
	  dataLabels: {
           enabled: false
        },
	  stroke: {
	    width: 0,
	  },
	  series: [{
	    name: 'Basic Digital Skills',
	    data: [basic]
	  }],
	  title: {
	    floating: true,
	    offsetX: -10,
	    offsetY: 5,
	    text: 'Basic Digital Skills'
	  },
	  subtitle: {
	    floating: true,
	    align: 'right',
	    offsetY: 0,
	    text: basic+'%',
	    style: {
	      fontSize: '20px'
	    }
	  },
	  fill: {
	    type: 'gradient',
	    gradient: {
	      gradientToColors: ['#6078ea']
	    },
	  },
	  colors: ['#17ead9'],
	  tooltip: {
	  	theme: 'dark',
	  	x: {
              show: false
             },
	    enabled: true
	    
	  },
	  xaxis: {
	    categories: ['Basic Digital Skills'],
	  },
	  yaxis: {
	    max: 100
	  },
	}

	var chartProgress2 = new ApexCharts(document.querySelector('#organic-search'), optionsProgress2);
	chartProgress2.render();




// chart 8

    var optionsProgress3 = {
	  chart: {
	    height: 70,
	    type: 'bar',
	    stacked: true,
	    sparkline: {
	      enabled: true
	    }
	  },
	  plotOptions: {
	    bar: {
	      horizontal: true,
	      barHeight: '20%',
	      colors: {
	        backgroundBarColors: ['#40475D']
	      }
	    },
	  },
	  dataLabels: {
           enabled: false
        },
	  stroke: {
	    width: 0,
	  },
	  series: [{
	    name: 'Intermediate Digital Skills',
	    data: [intemediate]
	  }],
	  title: {
	    floating: true,
	    offsetX: -10,
	    offsetY: 5,
	    text: 'Intermediate Digital Skills'
	  },
	  subtitle: {
	    floating: true,
	    align: 'right',
	    offsetY: 0,
	    text: intemediate+'%',
	    style: {
	      fontSize: '20px'
	    }
	  },
	  fill: {
	    type: 'gradient',
	    gradient: {
	      gradientToColors: ['#6094ea']
	    },
	  },
	  colors: ['#f02fc2'],
	  tooltip: {
	  	theme: 'dark',
	  	x: {
              show: false
             },
	    enabled: true
	    
	  },
	  xaxis: {
	    categories: ['Intermediate Digital Skills'],
	  },
	  yaxis: {
	    max: 100
	  },
	}

   var chartProgress3 = new ApexCharts(document.querySelector('#referral'), optionsProgress3);
	chartProgress3.render();
	
	
	
	
	// chart 9

    var optionsProgress4 = {
	  chart: {
	    height: 70,
	    type: 'bar',
	    stacked: true,
	    sparkline: {
	      enabled: true
	    }
	  },
	  plotOptions: {
	    bar: {
	      horizontal: true,
	      barHeight: '20%',
	      colors: {
	        backgroundBarColors: ['#40475D']
	      }
	    },
	  },
	  dataLabels: {
           enabled: false
        },
	  stroke: {
	    width: 0,
	  },
	  series: [{
	    name: 'Advanced digital skills',
	    data: [advanced]
	  }],
	  title: {
	    floating: true,
	    offsetX: -10,
	    offsetY: 5,
	    text: 'Advanced digital skills'
	  },
	  subtitle: {
	    floating: true,
	    align: 'right',
	    offsetY: 0,
	    text: advanced+'%',
	    style: {
	      fontSize: '20px'
	    }
	  },
	  fill: {
	    type: 'gradient',
	    gradient: {
	      gradientToColors: ['#6a11cb']
	    },
	  },
	  colors: ['#2575fc'],
	  tooltip: {
	  	theme: 'dark',
	  	x: {
              show: false
             },
	    enabled: true
	    
	  },
	  xaxis: {
	    categories: ['Advanced digital skills'],
	  },
	  yaxis: {
	    max: 100
	  },
	}

   var chartProgress4 = new ApexCharts(document.querySelector('#social'), optionsProgress4);
	chartProgress4.render();



    
   // chart 10

    var options = {
      chart: {
        height: 270,
        type: 'radialBar',
        toolbar: {
          show: false
        }
      },
      plotOptions: {
        radialBar: {
          //startAngle: -135,
          //endAngle: 225,
           hollow: {
            margin: 0,
            size: '85%',
            background: 'transparent',
            image: undefined,
            imageOffsetX: 0,
            imageOffsetY: 0,
            position: 'front',
            dropShadow: {
              enabled: true,
              top: 3,
              left: 0,
              blur: 4,
              color: 'rgba(0, 169, 255, 0.85)',
              opacity: 0.65
            }
          },
          track: {
            background: '#fff',
            strokeWidth: '67%',
            margin: 0, // margin is in pixels
            dropShadow: {
              enabled: true,
              top: -3,
              left: 0,
              blur: 4,
			  color: 'rgba(0, 169, 255, 0.85)',
              opacity: 0.65
            }
          },

          dataLabels: { 
            showOn: 'always',
            name: {
              offsetY: -5,
              show: false,
              color: '#000',
              fontSize: '14px'
            },
            value: {
              formatter: function (val) {
						return val + "%";
					},
              color: '#000',
              fontSize: '35px',
              show: true,
			  offsetY: 10,
            }
          }
        }
      },
      fill: {
        type: 'gradient',
        gradient: {
          shade: 'light',
          type: 'horizontal',
          shadeIntensity: 0.5,
          gradientToColors: ['#0072ff'],
          inverseColors: false,
          opacityFrom: 1,
          opacityTo: 1,
          stops: [0, 100]
        }
      },
      colors: ["#00c8ff"],
      series: [regist_rate],
      stroke: {
        lineCap: 'round',
        //dashArray: 4
      },
      labels: ['Newsletter Open Rate'],

    }

    var chart = new ApexCharts(
      document.querySelector("#newsletter-open-rate"),
      options
    );

    chart.render();

		
	// chart 11


	var options = {
      chart: {
        height: 270,
        type: 'radialBar',
        toolbar: {
          show: false
        }
      },
      plotOptions: {
        radialBar: {
         //startAngle: -135,
         // endAngle: 225,
           hollow: {
            margin: 0,
            size: '85%',
            background: 'transparent',
            image: undefined,
            imageOffsetX: 0,
            imageOffsetY: 0,
            position: 'front',
            dropShadow: {
              enabled: true,
              top: 3,
              left: 0,
              blur: 4,
              color: 'rgba(229, 124, 107, 0.85)',
              opacity: 0.65
            }
          },
          track: {
            background: '#fff',
            strokeWidth: '67%',
            margin: 0, // margin is in pixels
            dropShadow: {
              enabled: true,
              top: -3,
              left: 0,
              blur: 4,
			  color: 'rgba(229, 124, 107, 0.85)',
              opacity: 0.65
            }
          },

          dataLabels: { 
            showOn: 'always',
            name: {
              offsetY: -5,
              show: false,
              color: '#000',
              fontSize: '14px'
            },
            value: {
              formatter: function (val) {
						return val + "%";
					},
              color: '#000',
              fontSize: '35px',
              show: true,
			  offsetY: 10,
            }
          }
        }
      },
      fill: {
        type: 'gradient',
        gradient: {
          shade: 'light',
          type: 'horizontal',
          shadeIntensity: 0.5,
          gradientToColors: ['#db36a4'],
          inverseColors: false,
          opacityFrom: 1,
          opacityTo: 1,
          stops: [0, 100]
        }
      },
      colors: ["#f7ff00"],
      series: [certi_rate],
      stroke: {
        lineCap: 'round',
        //dashArray: 4
      },
      labels: ['Click Through Rate'],

    }

    var chart = new ApexCharts(
      document.querySelector("#click-through-rate"),
      options
    );

    chart.render();	











    });