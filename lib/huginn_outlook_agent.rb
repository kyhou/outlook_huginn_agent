require 'huginn_agent'

#HuginnAgent.load 'huginn_outlook_agent/concerns/my_agent_concern'
HuginnAgent.register 'huginn_outlook_agent/outlook_agent'
HuginnAgent.register 'huginn_outlook_agent/outlook_sender_agent'
