import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Check } from "lucide-react";
import { useState } from "react";

const SubscriptionPanel = () => {
  const [claudeApiKey, setClaudeApiKey] = useState("");

  const plans = [
    {
      name: "Starter",
      price: "$29",
      period: "/month",
      description: "Perfect for individual analysts and small teams",
      features: [
        "5 AI model generations per month",
        "Basic financial templates",
        "Excel integration",
        "Email support",
      ],
      popular: false,
    },
    {
      name: "Professional",
      price: "$79",
      period: "/month",
      description: "For growing teams and complex models",
      features: [
        "25 AI model generations per month",
        "Advanced financial templates",
        "Excel & Google Sheets integration",
        "Priority support",
        "Custom model training",
        "Team collaboration",
      ],
      popular: true,
    },
    {
      name: "Unlimited",
      price: "$199",
      period: "/month",
      description: "For large organizations with advanced needs",
      features: [
        "Unlimited AI model generations",
        "All financial templates",
        "Full platform integration",
        "24/7 dedicated support",
        "Custom model training",
        "Advanced team management",
        "API access",
        "On-premise deployment",
      ],
      popular: false,
    },
  ];

  return (
    <div className="max-w-2xl mx-auto">
      <div className="text-center mb-12">
        <h2 className="text-3xl font-bold text-gray-900 mb-4">
          Choose Your Plan
        </h2>
        <p className="text-gray-600 text-lg">
          Select the perfect plan for your financial modeling needs
        </p>
      </div>

      <div className="space-y-6">
        {plans.map((plan, index) => (
          <Card
            key={index}
            className={`relative ${
              plan.popular
                ? "border-red-600 shadow-lg"
                : "border-gray-200"
            }`}
          >
            {plan.popular && (
              <Badge className="absolute -top-3 left-1/2 transform -translate-x-1/2 bg-red-600 text-white">
                Most Popular
              </Badge>
            )}
            
            <CardHeader className="text-center pb-4">
              <CardTitle className="text-xl font-semibold text-gray-900">
                {plan.name}
              </CardTitle>
              <div className="mt-4">
                <span className="text-4xl font-bold text-gray-900">
                  {plan.price}
                </span>
                <span className="text-gray-600">{plan.period}</span>
              </div>
              <p className="text-gray-600 mt-2">{plan.description}</p>
            </CardHeader>

            <CardContent>
              <ul className="space-y-3 mb-8">
                {plan.features.map((feature, featureIndex) => (
                  <li key={featureIndex} className="flex items-center gap-3">
                    <Check size={16} className="text-green-600 flex-shrink-0" />
                    <span className="text-gray-700">{feature}</span>
                  </li>
                ))}
              </ul>

              {plan.name === "Unlimited" && (
                <div className="mb-6 space-y-2">
                  <Label htmlFor="claude-api-key" className="text-sm font-medium text-gray-700">
                    Claude API Key (Optional)
                  </Label>
                  <Input
                    id="claude-api-key"
                    type="password"
                    placeholder="Enter your Claude API key..."
                    value={claudeApiKey}
                    onChange={(e) => setClaudeApiKey(e.target.value)}
                    className="w-full"
                  />
                  <p className="text-xs text-gray-500">
                    Use your own Claude API key for enhanced performance and unlimited usage
                  </p>
                </div>
              )}

              <Button
                className={`w-full ${
                  plan.popular
                    ? "bg-red-600 hover:bg-red-700 text-white"
                    : "bg-black hover:bg-gray-800 text-white"
                }`}
              >
                Get Started
              </Button>
            </CardContent>
          </Card>
        ))}
      </div>

      <div className="mt-12 text-center">
        <p className="text-gray-600">
          All plans include a 14-day free trial. No credit card required.
        </p>
      </div>
    </div>
  );
};

export default SubscriptionPanel;
