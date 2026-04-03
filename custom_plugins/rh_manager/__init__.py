import random
from collections import defaultdict

def initialize(rhapi):

    APP = rhapi.server.APP
    request = rhapi.server.request


    # =========================
    # STATUS
    # =========================

    @APP.route('/rhapi/status', methods=['GET'])
    def status():
        return {"status": "running"}
