(function () {
  'use strict';

  angular.module('office365app')
    .controller('plannerController', ['$scope', '$window', 'dataService', plannerController]);

  /**
   * Controller constructor
   */
  function plannerController($scope, $window, dataService) {
    var vm = this;
    vm.loading = false;
    vm.myPlans = [];
    vm.navigate = navigate;

    activate();

    function activate() {
      loadUserData();
    }

    function navigate($event, url) {
      // only process the click if it not occured on a hyperlink
      var $target = jQuery($event.target).parent('a');
      if ($target.length === 0) {
        $window.open(url, '_blank');
      }
    }

    function loadUserData() {
      vm.loading = true;
      vm.myPlans.length = 0;

      dataService.getMyPlans().then(function (myPlans) {
        console && console.log(myPlans);

        myPlans.forEach(function (plan) {
          vm.myPlans.push(plan);
          loadPlanDetails(plan);
        });
      }, function (err) {
        console.error(err);
      }).finally(function () {
        vm.loading = false;
      });
    }

    function loadPlanDetails(plan) {
      dataService.getPlanDetails(plan.owner, plan.ownerOdataId).then(function (planDetails) {
        for (var i = 0; i < vm.myPlans.length; i++) {
          var g = vm.myPlans[i];
          if (g.id !== plan.id) {
            continue;
          }

          g.description = planDetails.description;
          g.unseenCount = planDetails.unseenCount;
          g.picture = planDetails.picture;
        }
      }, function (err) {
        console.error(err);
      });
    }
  }
    
})();